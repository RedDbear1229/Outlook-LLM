using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using MailPrioritizer.Models;
using MailPrioritizer.Utils;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MailPrioritizer.Services
{
    /// <summary>메일 분석 핵심 엔진. LLM 호출 → UserProperty 저장 → 폴더 이동.</summary>
    public class MailProcessor
    {
        private readonly LlmService _llmService;
        private readonly FolderManager _folderManager;
        private AppConfig _config;

        public MailProcessor(LlmService llmService, FolderManager folderManager, AppConfig config)
        {
            _llmService = llmService;
            _folderManager = folderManager;
            _config = config;
        }

        public void ReloadConfig(AppConfig newConfig)
        {
            _config = newConfig;
        }

        // ──────────────────────────────────────────────────────────────
        // 단건 분석
        // ──────────────────────────────────────────────────────────────

        /// <summary>
        /// 메일 1건을 분석한다. STA 스레드에서 호출해야 한다.
        /// await 후에도 ConfigureAwait(true) 기본값으로 STA 스레드 복귀가 보장된다.
        /// </summary>
        public async Task<MailAnalysis> AnalyzeSingleAsync(
            Outlook.MailItem mail,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            Logger.Info("AnalyzeSingleAsync: start, subject=\"" + (mail.Subject ?? "") + "\"");

            // 이미 분석된 메일이면 저장된 결과 반환
            var existing = LoadExistingAnalysis(mail);
            if (existing != null)
            {
                Logger.Info("AnalyzeSingleAsync: existing analysis found, priority=" + existing.Priority);
                return existing;
            }

            // COM 데이터를 STA 스레드에서 미리 추출
            string subject = mail.Subject ?? "";
            string body = mail.Body ?? "";
            string sender = GetSender(mail);

            // LLM 호출 (ThreadPool에서 실행, await 후 STA 복귀)
            var analysis = await _llmService.AnalyzeMailAsync(
                subject, body, sender, cancellationToken);

            // STA 스레드로 복귀 후 COM 작업
            if (_config.Display.TagSubjectWithPriority && !analysis.IsFallback)
                TagSubject(mail, analysis.Priority);

            SaveAnalysisToMail(mail, analysis);

            if (_config.Classification.AutoMoveToFolder && !analysis.IsFallback)
                _folderManager.MoveToFolder(mail, analysis.Priority);

            Logger.Info("AnalyzeSingleAsync: complete, priority=" + analysis.Priority);
            return analysis;
        }

        // ──────────────────────────────────────────────────────────────
        // 일괄 분석
        // ──────────────────────────────────────────────────────────────

        private const int ConsecutiveFailureThreshold = 5;

        public class BatchProgress
        {
            public int Total { get; set; }
            public int Processed { get { return Urgent + High + Normal + Low + Failed; } }
            public int Urgent { get; set; }
            public int High { get; set; }
            public int Normal { get; set; }
            public int Low { get; set; }
            public int Failed { get; set; }
            public string CurrentSubject { get; set; }
            public List<string> FailedSubjects { get; set; } = new List<string>();
        }

        /// <summary>
        /// 받은편지함의 미분석 메일을 모두 분석한다.
        /// progress 콜백은 STA 스레드에서 호출된다.
        /// </summary>
        public async Task<BatchProgress> AnalyzeInboxAsync(
            Outlook.Application outlookApp,
            IProgress<BatchProgress> progress,
            CancellationToken cancellationToken)
        {
            var result = new BatchProgress();

            // 미분석 메일 목록 수집 (EntryID로 참조 — 이동 후 MailItem 무효화 방지)
            var entryIds = CollectUnanalyzedEntryIds(outlookApp, cancellationToken);
            result.Total = entryIds.Count;
            Logger.Info("AnalyzeInboxAsync: start, total=" + entryIds.Count + " unanalyzed mails");
            progress?.Report(result);

            int consecutiveFailures = 0;
            // 배치 크기는 동시성 제한보다 크게 설정하여 LLM 파이프라인 효율 극대화
            int batchSize = Math.Max(_config.Processing.ConcurrentRequests * 5, 15);
            bool aborted = false;

            // 배치 단위로 병렬 처리: COM 읽기(STA) → LLM 병렬 호출 → COM 저장(STA)
            for (int i = 0; i < entryIds.Count && !aborted; i += batchSize)
            {
                cancellationToken.ThrowIfCancellationRequested();
                int count = Math.Min(batchSize, entryIds.Count - i);

                // Phase 1: STA에서 COM 데이터 추출
                var batchItems = new List<MailDataItem>();
                for (int j = 0; j < count; j++)
                {
                    string entryId = entryIds[i + j];
                    Outlook.MailItem mail = null;
                    try
                    {
                        mail = outlookApp.Session.GetItemFromID(entryId) as Outlook.MailItem;
                        if (mail == null) continue;

                        batchItems.Add(new MailDataItem
                        {
                            EntryId = entryId,
                            Subject = mail.Subject ?? "",
                            Body = mail.Body ?? "",
                            Sender = GetSender(mail)
                        });

                        result.CurrentSubject = mail.Subject ?? "(제목 없음)";
                        progress?.Report(result);
                    }
                    finally
                    {
                        ComHelper.Release(mail);
                    }
                }

                if (batchItems.Count == 0) continue;

                // Phase 2: LLM 호출 병렬 실행 (SemaphoreSlim이 동시성 제한)
                var tasks = new Task<MailAnalysis>[batchItems.Count];
                for (int j = 0; j < batchItems.Count; j++)
                {
                    var item = batchItems[j];
                    Logger.Debug("AnalyzeInboxAsync: processing \"" + item.Subject + "\"");
                    tasks[j] = _llmService.AnalyzeMailAsync(
                        item.Subject, item.Body, item.Sender, cancellationToken);
                }

                MailAnalysis[] analyses;
                try
                {
                    analyses = await Task.WhenAll(tasks);
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    // 전체 배치 실패 시 개별 결과 확인
                    Logger.Error("AnalyzeInboxAsync: batch LLM call error", ex);
                    analyses = new MailAnalysis[batchItems.Count];
                    for (int j = 0; j < tasks.Length; j++)
                    {
                        if (tasks[j].Status == TaskStatus.RanToCompletion)
                            analyses[j] = tasks[j].Result;
                        else
                            analyses[j] = MailAnalysis.CreateFallback("배치 호출 오류");
                    }
                }

                // Phase 3: STA에서 COM 작업 (UserProperty 저장 + 폴더 이동) — 순차 처리
                for (int j = 0; j < batchItems.Count; j++)
                {
                    var item = batchItems[j];
                    var analysis = analyses[j];

                    Outlook.MailItem freshMail = null;
                    try
                    {
                        freshMail = outlookApp.Session.GetItemFromID(item.EntryId) as Outlook.MailItem;
                        if (freshMail != null)
                        {
                            if (_config.Display.TagSubjectWithPriority && !analysis.IsFallback)
                                TagSubject(freshMail, analysis.Priority);

                            SaveAnalysisToMail(freshMail, analysis);

                            if (_config.Classification.AutoMoveToFolder && !analysis.IsFallback)
                                _folderManager.MoveToFolder(freshMail, analysis.Priority);
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Error("AnalyzeInboxAsync: failed to save/move mail", ex);
                    }
                    finally
                    {
                        ComHelper.Release(freshMail);
                    }

                    bool success = false;
                    if (!analysis.IsFallback)
                    {
                        switch (analysis.Priority)
                        {
                            case Priority.Urgent: result.Urgent++; break;
                            case Priority.High:   result.High++;   break;
                            case Priority.Normal: result.Normal++; break;
                            case Priority.Low:    result.Low++;    break;
                        }
                        success = true;
                    }
                    else
                    {
                        result.Failed++;
                        result.FailedSubjects.Add(item.Subject);
                    }

                    // 연속 실패 감지
                    if (success)
                    {
                        consecutiveFailures = 0;
                    }
                    else
                    {
                        consecutiveFailures++;
                        if (consecutiveFailures >= ConsecutiveFailureThreshold)
                        {
                            Logger.Warn("AnalyzeInboxAsync: " + consecutiveFailures + " consecutive failures");
                            var answer = System.Windows.Forms.MessageBox.Show(
                                string.Format("{0}건 연속 실패했습니다.\nAPI 연결 상태를 확인해 주세요.\n\n계속 진행하시겠습니까?",
                                    consecutiveFailures),
                                "MailPrioritizer",
                                System.Windows.Forms.MessageBoxButtons.YesNo,
                                System.Windows.Forms.MessageBoxIcon.Warning);
                            if (answer == System.Windows.Forms.DialogResult.No)
                            {
                                aborted = true;
                                break;
                            }
                            consecutiveFailures = 0;
                        }
                    }

                    progress?.Report(result);
                }
            }

            Logger.Info(string.Format(
                "AnalyzeInboxAsync: batch complete. Processed={0}, Failed={1}, Urgent={2}, High={3}, Normal={4}, Low={5}",
                result.Processed, result.Failed, result.Urgent, result.High, result.Normal, result.Low));

            return result;
        }

        /// <summary>배치 처리 시 COM에서 추출한 메일 데이터.</summary>
        private class MailDataItem
        {
            public string EntryId;
            public string Subject;
            public string Body;
            public string Sender;
        }

        // ──────────────────────────────────────────────────────────────
        // UserProperty 저장
        // ──────────────────────────────────────────────────────────────

        public void SaveAnalysisToMail(Outlook.MailItem mail, MailAnalysis analysis)
        {
            Outlook.UserProperties props = null;
            Outlook.UserProperty propPriority  = null;
            Outlook.UserProperty propSummary   = null;
            Outlook.UserProperty propReason    = null;
            Outlook.UserProperty propAnalyzed  = null;

            try
            {
                props = mail.UserProperties;

                propPriority = props.Find(MailPropertyNames.Priority) ??
                               props.Add(MailPropertyNames.Priority, Outlook.OlUserPropertyType.olText);
                propPriority.Value = analysis.Priority.ToString();

                propSummary = props.Find(MailPropertyNames.Summary) ??
                              props.Add(MailPropertyNames.Summary, Outlook.OlUserPropertyType.olText);
                propSummary.Value = analysis.Summary ?? "";

                propReason = props.Find(MailPropertyNames.PriorityReason) ??
                             props.Add(MailPropertyNames.PriorityReason, Outlook.OlUserPropertyType.olText);
                propReason.Value = analysis.PriorityReason ?? "";

                propAnalyzed = props.Find(MailPropertyNames.Analyzed) ??
                               props.Add(MailPropertyNames.Analyzed, Outlook.OlUserPropertyType.olYesNo);
                propAnalyzed.Value = true;

                mail.Save();
            }
            finally
            {
                ComHelper.ReleaseAll(propAnalyzed, propReason, propSummary, propPriority, props);
            }
        }

        // ──────────────────────────────────────────────────────────────
        // UserProperty 읽기
        // ──────────────────────────────────────────────────────────────

        /// <summary>
        /// 메일에 저장된 기존 분석 결과를 읽어 반환.
        /// 분석 기록이 없으면 null 반환.
        /// </summary>
        public MailAnalysis LoadExistingAnalysis(Outlook.MailItem mail)
        {
            Outlook.UserProperties props = null;
            Outlook.UserProperty propAnalyzed = null;
            Outlook.UserProperty propPriority = null;
            Outlook.UserProperty propSummary  = null;
            Outlook.UserProperty propReason   = null;

            try
            {
                props = mail.UserProperties;

                propAnalyzed = props.Find(MailPropertyNames.Analyzed);
                if (propAnalyzed == null || !IsAnalyzed(propAnalyzed.Value))
                    return null;

                propPriority = props.Find(MailPropertyNames.Priority);
                propSummary  = props.Find(MailPropertyNames.Summary);
                propReason   = props.Find(MailPropertyNames.PriorityReason);

                return new MailAnalysis
                {
                    Priority      = PriorityExtensions.FromString(propPriority?.Value?.ToString()),
                    Summary       = propSummary?.Value?.ToString() ?? "",
                    PriorityReason = propReason?.Value?.ToString() ?? "",
                    AnalyzedAt    = DateTime.Now,
                    IsFallback    = false
                };
            }
            finally
            {
                ComHelper.ReleaseAll(propReason, propSummary, propPriority, propAnalyzed, props);
            }
        }

        // ──────────────────────────────────────────────────────────────
        // 헬퍼
        // ──────────────────────────────────────────────────────────────

        /// <summary>단일 메일의 LLM_Analyzed 플래그를 초기화한다.</summary>
        public static void ClearAnalysisFlag(Outlook.MailItem mail)
        {
            Outlook.UserProperties props = null;
            Outlook.UserProperty prop = null;
            try
            {
                props = mail.UserProperties;
                prop = props.Find(MailPropertyNames.Analyzed);
                if (prop != null)
                {
                    prop.Value = false;
                    mail.Save();
                }
            }
            finally
            {
                ComHelper.ReleaseAll(prop, props);
            }
        }

        /// <summary>UserProperty 값이 "분석됨"인지 확인. olYesNo 타입은 bool 또는 int로 저장될 수 있다.</summary>
        internal static bool IsAnalyzed(object value)
        {
            if (value is bool b) return b;
            if (value is int i) return i != 0;
            return false;
        }

        private static string GetSender(Outlook.MailItem mail)
        {
            return mail.SenderEmailAddress ?? mail.SenderName ?? "";
        }

        /// <summary>제목에 우선순위 태그 삽입. mail.Save()는 호출하지 않음 — 호출자가 일괄 Save.</summary>
        private void TagSubject(Outlook.MailItem mail, Priority priority)
        {
            string tag = "[" + priority.ToKorean() + "] ";
            if (mail.Subject != null && !mail.Subject.StartsWith("["))
            {
                mail.Subject = tag + mail.Subject;
            }
        }

        /// <summary>DASL 필터로 UserProperty 기반 검색에 사용하는 프로퍼티 태그.</summary>
        private const string DaslAnalyzedProp =
            "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/" + MailPropertyNames.Analyzed;

        private List<string> CollectUnanalyzedEntryIds(Outlook.Application app, CancellationToken ct)
        {
            var list = new List<string>();
            Outlook.MAPIFolder inbox = null;
            Outlook.Items allItems = null;
            Outlook.Items analyzedItems = null;

            try
            {
                inbox = _folderManager.GetTargetInbox();
                allItems = inbox.Items;

                // DASL 필터로 이미 분석된 메일의 EntryID 집합을 먼저 수집
                var analyzedIds = new HashSet<string>();
                try
                {
                    string filter = "@SQL=\"" + DaslAnalyzedProp + "\" = 1";
                    analyzedItems = allItems.Restrict(filter);

                    foreach (object item in analyzedItems)
                    {
                        ct.ThrowIfCancellationRequested();
                        var mail = item as Outlook.MailItem;
                        if (mail != null)
                            analyzedIds.Add(mail.EntryID);
                        ComHelper.Release(item);
                    }
                }
                catch (Exception ex)
                {
                    // DASL 필터 실패 시 폴백: 전체 순회
                    Logger.Warn("CollectUnanalyzedEntryIds: DASL filter failed, falling back to full scan. " + ex.Message);
                    ComHelper.Release(analyzedItems);
                    analyzedItems = null;
                    return CollectUnanalyzedEntryIdsFallback(allItems, ct);
                }

                // 전체 메일 중 분석되지 않은 것만 수집 (MailItem 여부만 확인)
                foreach (object item in allItems)
                {
                    ct.ThrowIfCancellationRequested();
                    var mail = item as Outlook.MailItem;
                    if (mail != null && !analyzedIds.Contains(mail.EntryID))
                        list.Add(mail.EntryID);
                    ComHelper.Release(item);
                }
            }
            finally
            {
                ComHelper.ReleaseAll(analyzedItems, allItems, inbox);
            }

            Logger.Info("CollectUnanalyzedEntryIds: found " + list.Count + " unanalyzed mails");
            return list;
        }

        /// <summary>DASL 필터 실패 시 폴백: UserProperty를 직접 읽어 확인.</summary>
        private List<string> CollectUnanalyzedEntryIdsFallback(Outlook.Items items, CancellationToken ct)
        {
            var list = new List<string>();
            foreach (object item in items)
            {
                ct.ThrowIfCancellationRequested();
                var mail = item as Outlook.MailItem;
                if (mail == null) { ComHelper.Release(item); continue; }

                Outlook.UserProperties props = null;
                Outlook.UserProperty prop = null;
                try
                {
                    props = mail.UserProperties;
                    prop = props.Find(MailPropertyNames.Analyzed);
                    if (prop == null || !IsAnalyzed(prop.Value))
                        list.Add(mail.EntryID);
                }
                finally
                {
                    ComHelper.ReleaseAll(prop, props, mail);
                }
            }
            return list;
        }

        // ──────────────────────────────────────────────────────────────
        // 받은편지함 통계 (#9)
        // ──────────────────────────────────────────────────────────────

        public class InboxStats
        {
            public int Total;
            public int Analyzed;
            public int Urgent;
            public int High;
            public int Normal;
            public int Low;
        }

        /// <summary>받은편지함의 분석 통계를 수집한다.</summary>
        public InboxStats CollectInboxStats(Outlook.Application app)
        {
            var stats = new InboxStats();
            Outlook.MAPIFolder inbox = null;
            Outlook.Items items = null;

            try
            {
                inbox = _folderManager.GetTargetInbox();
                items = inbox.Items;

                foreach (object item in items)
                {
                    var mail = item as Outlook.MailItem;
                    if (mail == null)
                    {
                        ComHelper.Release(item);
                        continue;
                    }

                    stats.Total++;
                    Outlook.UserProperties props = null;
                    Outlook.UserProperty propAnalyzed = null;
                    Outlook.UserProperty propPriority = null;
                    try
                    {
                        props = mail.UserProperties;
                        propAnalyzed = props.Find(MailPropertyNames.Analyzed);
                        if (propAnalyzed != null && IsAnalyzed(propAnalyzed.Value))
                        {
                            stats.Analyzed++;
                            propPriority = props.Find(MailPropertyNames.Priority);
                            switch (PriorityExtensions.FromString(propPriority?.Value?.ToString()))
                            {
                                case Priority.Urgent: stats.Urgent++; break;
                                case Priority.High:   stats.High++;   break;
                                case Priority.Normal: stats.Normal++; break;
                                case Priority.Low:    stats.Low++;    break;
                            }
                        }
                    }
                    finally
                    {
                        ComHelper.ReleaseAll(propPriority, propAnalyzed, props, mail);
                    }
                }
            }
            finally
            {
                ComHelper.ReleaseAll(items, inbox);
            }

            return stats;
        }

        // ──────────────────────────────────────────────────────────────
        // 전체 재분석 (#7)
        // ──────────────────────────────────────────────────────────────

        /// <summary>받은편지함의 모든 메일에서 LLM_Analyzed 플래그를 초기화한다.</summary>
        public int ResetAllAnalysisFlags(Outlook.Application app)
        {
            int count = 0;
            Outlook.MAPIFolder inbox = null;
            Outlook.Items items = null;

            try
            {
                inbox = _folderManager.GetTargetInbox();
                items = inbox.Items;

                foreach (object item in items)
                {
                    var mail = item as Outlook.MailItem;
                    if (mail == null)
                    {
                        ComHelper.Release(item);
                        continue;
                    }

                    try
                    {
                        var existing = LoadExistingAnalysis(mail);
                        if (existing != null)
                        {
                            ClearAnalysisFlag(mail);
                            count++;
                        }
                    }
                    finally
                    {
                        ComHelper.Release(mail);
                    }
                }
            }
            finally
            {
                ComHelper.ReleaseAll(items, inbox);
            }

            Logger.Info("ResetAllAnalysisFlags: reset " + count + " mails");
            return count;
        }

        // ──────────────────────────────────────────────────────────────
        // 분석 결과 내보내기 (#12)
        // ──────────────────────────────────────────────────────────────

        /// <summary>분석된 메일의 결과를 CSV로 내보낸다. 반환값: 내보낸 건수.</summary>
        public int ExportAnalyzedMails(Outlook.Application app, string filePath)
        {
            int count = 0;
            Outlook.MAPIFolder inbox = null;
            Outlook.Items items = null;

            var sb = new StringBuilder();
            sb.AppendLine("Subject,Sender,Priority,Summary,PriorityReason");

            try
            {
                inbox = _folderManager.GetTargetInbox();
                items = inbox.Items;

                foreach (object item in items)
                {
                    var mail = item as Outlook.MailItem;
                    if (mail == null)
                    {
                        ComHelper.Release(item);
                        continue;
                    }

                    Outlook.UserProperties props = null;
                    Outlook.UserProperty propAnalyzed = null;
                    Outlook.UserProperty propPriority = null;
                    Outlook.UserProperty propSummary = null;
                    Outlook.UserProperty propReason = null;
                    try
                    {
                        props = mail.UserProperties;
                        propAnalyzed = props.Find(MailPropertyNames.Analyzed);
                        if (propAnalyzed == null || !IsAnalyzed(propAnalyzed.Value))
                            continue;

                        propPriority = props.Find(MailPropertyNames.Priority);
                        propSummary = props.Find(MailPropertyNames.Summary);
                        propReason = props.Find(MailPropertyNames.PriorityReason);

                        sb.AppendLine(string.Format("{0},{1},{2},{3},{4}",
                            CsvEscape(mail.Subject ?? ""),
                            CsvEscape(GetSender(mail)),
                            CsvEscape(propPriority?.Value?.ToString() ?? ""),
                            CsvEscape(propSummary?.Value?.ToString() ?? ""),
                            CsvEscape(propReason?.Value?.ToString() ?? "")));
                        count++;
                    }
                    finally
                    {
                        ComHelper.ReleaseAll(propReason, propSummary, propPriority, propAnalyzed, props, mail);
                    }
                }
            }
            finally
            {
                ComHelper.ReleaseAll(items, inbox);
            }

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
            Logger.Info("ExportAnalyzedMails: exported " + count + " mails to " + filePath);
            return count;
        }

        private static string CsvEscape(string value)
        {
            if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
                return "\"" + value.Replace("\"", "\"\"") + "\"";
            return value;
        }
    }
}
