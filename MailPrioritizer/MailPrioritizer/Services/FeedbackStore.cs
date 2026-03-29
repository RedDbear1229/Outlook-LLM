using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MailPrioritizer.Models;
using MailPrioritizer.Utils;
using Newtonsoft.Json;

namespace MailPrioritizer.Services
{
    /// <summary>
    /// 사용자가 LLM 분류 결과를 수동으로 변경한 이력을 저장하고 정확도 통계를 제공한다.
    /// %AppData%/MailPrioritizer/feedback.json 에 영속 저장된다.
    /// </summary>
    public class FeedbackStore
    {
        private static readonly string FeedbackPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "MailPrioritizer", "feedback.json");

        private List<FeedbackItem> _items;
        private static readonly object _lock = new object();

        public class FeedbackItem
        {
            public string EntryId { get; set; }
            public string Subject { get; set; }
            public Priority LlmPriority { get; set; }
            public Priority UserPriority { get; set; }
            public DateTime ChangedAt { get; set; }
        }

        public class AccuracyStats
        {
            public int AnalyzedTotal { get; set; }
            public int CorrectionCount { get; set; }
            public double AccuracyPercent
            {
                get
                {
                    if (AnalyzedTotal == 0) return 0;
                    int correct = Math.Max(0, AnalyzedTotal - CorrectionCount);
                    return Math.Round(correct * 100.0 / AnalyzedTotal, 1);
                }
            }
            public string MostMisclassified { get; set; }
        }

        public FeedbackStore()
        {
            _items = LoadFromDisk();
        }

        public int CorrectionCount
        {
            get { lock (_lock) return _items.Count; }
        }

        /// <summary>
        /// 사용자가 우선순위를 변경했을 때 호출한다.
        /// LLM 결과와 동일하면 기록하지 않는다.
        /// </summary>
        public void Record(string entryId, string subject, Priority llmPriority, Priority userPriority)
        {
            if (llmPriority == userPriority) return;

            lock (_lock)
            {
                _items.RemoveAll(x => x.EntryId == entryId);
                _items.Add(new FeedbackItem
                {
                    EntryId = entryId,
                    Subject = subject ?? "",
                    LlmPriority = llmPriority,
                    UserPriority = userPriority,
                    ChangedAt = DateTime.Now
                });
                SaveToDisk();
            }
            Logger.Info(string.Format("FeedbackStore: {0}→{1} recorded for \"{2}\"",
                llmPriority, userPriority, subject));
        }

        /// <summary>분석 총 건수를 기준으로 정확도 통계를 계산한다.</summary>
        public AccuracyStats GetStats(int analyzedTotal)
        {
            lock (_lock)
            {
                string mostMisclassified = "";
                var topPattern = _items
                    .GroupBy(x => x.LlmPriority.ToKorean() + " → " + x.UserPriority.ToKorean())
                    .OrderByDescending(g => g.Count())
                    .FirstOrDefault();
                if (topPattern != null)
                    mostMisclassified = topPattern.Key + " (" + topPattern.Count() + "건)";

                return new AccuracyStats
                {
                    AnalyzedTotal = analyzedTotal,
                    CorrectionCount = _items.Count,
                    MostMisclassified = mostMisclassified
                };
            }
        }

        private List<FeedbackItem> LoadFromDisk()
        {
            try
            {
                if (!File.Exists(FeedbackPath)) return new List<FeedbackItem>();
                string json = File.ReadAllText(FeedbackPath);
                return JsonConvert.DeserializeObject<List<FeedbackItem>>(json) ?? new List<FeedbackItem>();
            }
            catch (Exception ex)
            {
                Logger.Warn("FeedbackStore.LoadFromDisk failed: " + ex.Message);
                return new List<FeedbackItem>();
            }
        }

        private void SaveToDisk()
        {
            try
            {
                string json = JsonConvert.SerializeObject(_items, Formatting.Indented);
                File.WriteAllText(FeedbackPath, json);
            }
            catch (Exception ex)
            {
                Logger.Warn("FeedbackStore.SaveToDisk failed: " + ex.Message);
            }
        }
    }
}
