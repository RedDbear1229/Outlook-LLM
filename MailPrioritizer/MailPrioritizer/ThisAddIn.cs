using System;
using System.Threading;
using System.Windows.Forms;
using MailPrioritizer.Config;
using MailPrioritizer.Models;
using MailPrioritizer.Services;
using MailPrioritizer.TaskPane;
using MailPrioritizer.Utils;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MailPrioritizer
{
    public partial class ThisAddIn
    {
        // ── 서비스 (ThisAddIn이 소유, Ribbon/Form에서 Globals.ThisAddIn으로 접근) ──
        internal ConfigManager ConfigManager { get; private set; }
        internal AppConfig Config { get; private set; }
        internal LlmService LlmService { get; private set; }
        internal FolderManager FolderManager { get; private set; }
        internal MailProcessor MailProcessor { get; private set; }
        internal RetryQueue RetryQueue { get; private set; }
        internal FeedbackStore FeedbackStore { get; private set; }

        // ── Task Pane ──
        private Microsoft.Office.Tools.CustomTaskPane _summaryPane;
        internal SummaryControl SummaryControl { get; private set; }

        // ── Explorer 이벤트 참조 (GC 방지) ──
        private Outlook.Explorer _explorer;

        // ── SelectionChange 디바운스 ──
        private System.Windows.Forms.Timer _selectionDebounceTimer;
        private string _lastSelectedEntryId;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // WinForms SynchronizationContext 보장 (STA 스레드 복귀를 위해 필수)
            if (SynchronizationContext.Current == null)
                SynchronizationContext.SetSynchronizationContext(
                    new WindowsFormsSynchronizationContext());

            // 로거 초기화 (모든 서비스보다 먼저)
            Logger.Initialize();

            // 서비스 초기화
            ConfigManager = new ConfigManager();
            Config        = ConfigManager.Load();
            Logger.Info("Startup: config loaded, endpoint=" + Config.Llm.Endpoint
                        + ", model=" + Config.Llm.ModelName);

            LlmService    = new LlmService(Config);
            FolderManager = new FolderManager(Application, Config);
            RetryQueue    = new RetryQueue();
            FeedbackStore = new FeedbackStore();
            MailProcessor = new MailProcessor(LlmService, FolderManager, Config, RetryQueue);
            Logger.Info("Startup: services created");

            // Custom Task Pane (우측 패널)
            SummaryControl = new SummaryControl();
            _summaryPane = CustomTaskPanes.Add(SummaryControl, "메일 분석");
            _summaryPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;

            // R-03: 마지막 상태 복원
            _summaryPane.Width   = Config.Display.TaskPaneWidth > 0
                ? Config.Display.TaskPaneWidth : 320;
            _summaryPane.Visible = Config.Display.TaskPaneVisible;

            // R-03: 상태 변경 시 자동 저장
            _summaryPane.VisibleChanged += (s, ev) => SaveTaskPaneState();

            // SelectionChange 디바운스 타이머 (200ms)
            _selectionDebounceTimer = new System.Windows.Forms.Timer { Interval = 200 };
            _selectionDebounceTimer.Tick += (s, ev) =>
            {
                _selectionDebounceTimer.Stop();
                Explorer_SelectionChange_Debounced();
            };

            // Explorer SelectionChange 이벤트
            _explorer = Application.ActiveExplorer();
            if (_explorer != null)
                _explorer.SelectionChange += Explorer_SelectionChange;

            // NewMailEx 이벤트 (#11 신규 메일 자동 분석)
            Application.NewMailEx += Application_NewMailEx;

            // R-02: 이전에 실패한 메일 재시도 (비동기, 시작을 블록하지 않음)
            if (RetryQueue.Count > 0)
            {
                Logger.Info("Startup: processing " + RetryQueue.Count + " items in retry queue");
                ProcessRetryQueueAsync();
            }

            Logger.Info("Startup complete");
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Logger.Info("Shutdown: begin");

            // 타이머 해제
            if (_selectionDebounceTimer != null)
            {
                _selectionDebounceTimer.Stop();
                _selectionDebounceTimer.Dispose();
                _selectionDebounceTimer = null;
            }

            // 이벤트 해제
            try { Application.NewMailEx -= Application_NewMailEx; }
            catch (Exception ex) { Logger.Error("Shutdown: failed to unsubscribe NewMailEx", ex); }

            if (_explorer != null)
            {
                try { _explorer.SelectionChange -= Explorer_SelectionChange; }
                catch (Exception ex) { Logger.Error("Shutdown: failed to unsubscribe Explorer event", ex); }
                ComHelper.Release(_explorer);
                _explorer = null;
            }

            // Task Pane 해제
            if (_summaryPane != null)
            {
                try { _summaryPane.Dispose(); }
                catch (Exception ex) { Logger.Error("Shutdown: failed to dispose TaskPane", ex); }
                _summaryPane = null;
            }

            // LlmService 해제 (HttpClient, SemaphoreSlim)
            if (LlmService != null)
            {
                try { LlmService.Dispose(); }
                catch (Exception ex) { Logger.Error("Shutdown: failed to dispose LlmService", ex); }
            }

            Logger.Info("Shutdown: complete");
        }

        // ──────────────────────────────────────────────────────────────
        // R-03: Task Pane 상태 저장
        // ──────────────────────────────────────────────────────────────

        private void SaveTaskPaneState()
        {
            try
            {
                Config.Display.TaskPaneVisible = _summaryPane.Visible;
                Config.Display.TaskPaneWidth   = _summaryPane.Width;
                ConfigManager.Save(Config);
            }
            catch (Exception ex)
            {
                Logger.Error("SaveTaskPaneState failed", ex);
            }
        }

        // ──────────────────────────────────────────────────────────────
        // R-02: 재시도 큐 처리
        // ──────────────────────────────────────────────────────────────

        private async void ProcessRetryQueueAsync()
        {
            var items = RetryQueue.GetAll();
            foreach (var item in items)
            {
                if (item.RetryCount >= RetryQueue.MaxRetries)
                {
                    RetryQueue.Remove(item.EntryId);
                    continue;
                }

                Outlook.MailItem mail = null;
                try
                {
                    mail = Application.Session.GetItemFromID(item.EntryId) as Outlook.MailItem;
                    if (mail == null)
                    {
                        RetryQueue.Remove(item.EntryId);
                        continue;
                    }

                    // 이미 성공 분석된 경우 큐에서 제거
                    var existing = MailProcessor.LoadExistingAnalysis(mail);
                    if (existing != null && !existing.IsFallback)
                    {
                        RetryQueue.Remove(item.EntryId);
                        continue;
                    }

                    // 이전 실패 결과(플래그)를 지우고 재분석
                    Services.MailProcessor.ClearAnalysisFlag(mail);
                    var analysis = await MailProcessor.AnalyzeSingleAsync(mail);
                    if (!analysis.IsFallback)
                        Logger.Info("ProcessRetryQueue: retry succeeded for entryId=" + item.EntryId);
                    else
                        RetryQueue.IncrementRetry(item.EntryId);
                }
                catch (Exception ex)
                {
                    Logger.Error("ProcessRetryQueue: failed for entryId=" + item.EntryId, ex);
                    RetryQueue.IncrementRetry(item.EntryId);
                }
                finally
                {
                    ComHelper.Release(mail);
                }
            }
        }

        // ──────────────────────────────────────────────────────────────
        // Explorer SelectionChange
        // ──────────────────────────────────────────────────────────────

        private void Explorer_SelectionChange()
        {
            // 빠른 키보드 탐색 시 마지막 선택만 처리하도록 디바운스
            _selectionDebounceTimer.Stop();
            _selectionDebounceTimer.Start();
        }

        private void Explorer_SelectionChange_Debounced()
        {
            if (SummaryControl == null || SummaryControl.IsDisposed) return;

            Outlook.Selection selection = null;
            Outlook.MailItem mail = null;
            try
            {
                if (_explorer == null) return;
                selection = _explorer.Selection;

                if (selection.Count == 1)
                {
                    mail = selection[1] as Outlook.MailItem;
                    if (mail != null)
                    {
                        // 같은 메일 재선택 시 불필요한 COM 작업 방지
                        if (mail.EntryID == _lastSelectedEntryId) return;
                        _lastSelectedEntryId = mail.EntryID;

                        var existing = MailProcessor.LoadExistingAnalysis(mail);
                        if (existing != null)
                            SummaryControl.DisplayAnalysis(mail, existing);
                        else
                            SummaryControl.ShowUnanalyzed(mail);
                    }
                    else
                    {
                        _lastSelectedEntryId = null;
                        SummaryControl.ShowNoSelection();
                    }
                }
                else
                {
                    _lastSelectedEntryId = null;
                    SummaryControl.ShowNoSelection();
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Explorer_SelectionChange: unhandled exception", ex);
            }
            finally
            {
                ComHelper.ReleaseAll(mail, selection);
            }
        }

        // ──────────────────────────────────────────────────────────────
        // NewMailEx — 신규 메일 자동 분석 (#11)
        // ──────────────────────────────────────────────────────────────

        private async void Application_NewMailEx(string entryIdCollection)
        {
            if (!Config.Processing.AutoAnalyzeNewMail) return;
            if (!IsConfigured()) return;

            string[] entryIds = entryIdCollection.Split(',');

            // 각 메일을 병렬 분석 (LlmService의 SemaphoreSlim이 동시성 제한)
            var tasks = new System.Collections.Generic.List<System.Threading.Tasks.Task>();
            foreach (string rawId in entryIds)
            {
                string entryId = rawId.Trim();
                if (string.IsNullOrEmpty(entryId)) continue;
                tasks.Add(AnalyzeNewMailAsync(entryId));
            }

            try
            {
                await System.Threading.Tasks.Task.WhenAll(tasks);
            }
            catch (Exception ex)
            {
                Logger.Error("NewMailEx: one or more auto-analyses failed", ex);
            }
        }

        private async System.Threading.Tasks.Task AnalyzeNewMailAsync(string entryId)
        {
            Outlook.MailItem mail = null;
            Outlook.NameSpace session = null;
            try
            {
                session = Application.Session;
                mail = session.GetItemFromID(entryId) as Outlook.MailItem;
                if (mail == null) return;

                Logger.Debug("NewMailEx: auto-analyzing entryId=" + entryId);
                var analysis = await MailProcessor.AnalyzeSingleAsync(mail);

                // 현재 선택 중인 메일이면 Task Pane 갱신
                RefreshTaskPaneIfSelected(entryId, analysis);
            }
            catch (Exception ex)
            {
                Logger.Error("NewMailEx: auto-analysis failed for entryId=" + entryId, ex);
            }
            finally
            {
                ComHelper.ReleaseAll(mail, session);
            }
        }

        private void RefreshTaskPaneIfSelected(string entryId, Models.MailAnalysis analysis)
        {
            if (SummaryControl == null || SummaryControl.IsDisposed) return;

            Outlook.Selection selection = null;
            Outlook.MailItem selected = null;
            try
            {
                if (_explorer == null) return;
                selection = _explorer.Selection;
                if (selection.Count != 1) return;

                selected = selection[1] as Outlook.MailItem;
                if (selected != null && selected.EntryID == entryId)
                    SummaryControl.DisplayAnalysis(selected, analysis);
            }
            finally
            {
                ComHelper.ReleaseAll(selected, selection);
            }
        }

        // ──────────────────────────────────────────────────────────────
        // 설정 변경 적용
        // ──────────────────────────────────────────────────────────────

        /// <summary>SettingsForm 저장 후 호출. 모든 서비스에 새 설정 반영.</summary>
        internal void ApplyConfigChange(AppConfig newConfig, string plainToken)
        {
            Logger.Info("ApplyConfigChange: applying new configuration");

            // plainToken이 비어 있지 않으면 암호화하여 저장
            ConfigManager.Save(newConfig, plainToken);

            // 서비스에 반영 (HttpClient 재생성 포함)
            // ApiToken은 메모리 내 항상 평문
            newConfig.Llm.ApiToken = plainToken;

            Config = newConfig;
            LlmService.ReloadConfig(newConfig);
            FolderManager.ReloadConfig(newConfig);
            MailProcessor.ReloadConfig(newConfig);

            // R-03: 설정 변경 시 Task Pane 너비도 반영
            if (_summaryPane != null && newConfig.Display.TaskPaneWidth > 0)
                _summaryPane.Width = newConfig.Display.TaskPaneWidth;

            Logger.Info("ApplyConfigChange: autoAnalyze=" + newConfig.Processing.AutoAnalyzeNewMail
                + ", targetStore=" + (string.IsNullOrEmpty(newConfig.Processing.TargetStoreId) ? "(default)" : "custom"));
        }

        // ──────────────────────────────────────────────────────────────
        // 설정 완료 여부 확인
        // ──────────────────────────────────────────────────────────────

        internal bool IsConfigured()
        {
            return !string.IsNullOrWhiteSpace(Config.Llm.Endpoint)
                && !string.IsNullOrWhiteSpace(Config.Llm.ApiToken)
                && !string.IsNullOrWhiteSpace(Config.Llm.ModelName);
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup  += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
