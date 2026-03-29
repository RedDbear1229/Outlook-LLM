using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using MailPrioritizer.Models;
using MailPrioritizer.Utils;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MailPrioritizer.TaskPane
{
    /// <summary>
    /// Custom Task Pane에 표시되는 메일 분석 결과 패널.
    /// Explorer의 SelectionChange 이벤트로 메일 전환 시 자동 갱신.
    /// </summary>
    public class SummaryControl : UserControl
    {
        private Panel _pnlPriorityBadge;
        private Label _lblPriorityText;
        private Label _lblSummaryHeader;
        private Label _lblSummary;
        private Label _lblReasonHeader;
        private Label _lblReason;
        private Panel _pnlDivider;
        private Label _lblChangePriority;
        private ComboBox _cmbPriority;
        private Button _btnReanalyze;
        private Button _btnMoveFolder;
        private Panel _pnlNotAnalyzed;
        private Label _lblNotAnalyzed;
        private Button _btnAnalyzeNow;
        private Panel _pnlAnalyzing;
        private Label _lblAnalyzing;
        private Panel _pnlStats;
        private Label _lblStatsTitle;
        private Label _lblStatsContent;
        private Button _btnRefreshStats;

        // 현재 표시 중인 메일 (EntryID로 참조, COM 객체 직접 보관 지양)
        private string _currentEntryId;

        public SummaryControl()
        {
            InitializeComponent();
            ShowNoSelection();
        }

        // ──────────────────────────────────────────────────────────────
        // 외부 호출 API
        // ──────────────────────────────────────────────────────────────

        public void ShowAnalyzing()
        {
            SafeInvoke(() => SetViewState(showAnalyzing: true));
        }

        public void DisplayAnalysis(Outlook.MailItem mail, MailAnalysis analysis)
        {
            SafeInvoke(() =>
            {
                _currentEntryId = mail?.EntryID;
                RenderAnalysis(analysis);
            });
        }

        public void ShowNoSelection()
        {
            SafeInvoke(() =>
            {
                SetViewState(showNotAnalyzed: true);
                _lblNotAnalyzed.Text = "메일을 선택하면 분석 결과가 표시됩니다.";
                _btnAnalyzeNow.Visible = false;
            });
        }

        public void ShowUnanalyzed(Outlook.MailItem mail)
        {
            SafeInvoke(() =>
            {
                _currentEntryId = mail?.EntryID;
                SetViewState(showNotAnalyzed: true);
                _lblNotAnalyzed.Text = "아직 분석되지 않은 메일입니다.";
                _btnAnalyzeNow.Visible = true;
            });
        }

        public void ShowError(string message)
        {
            SafeInvoke(() =>
            {
                SetViewState(showNotAnalyzed: true);
                _lblNotAnalyzed.Text = "오류: " + message;
                _btnAnalyzeNow.Visible = false;
            });
        }

        // ──────────────────────────────────────────────────────────────
        // 렌더링
        // ──────────────────────────────────────────────────────────────

        /// <summary>모든 가시성 플래그를 한번에 설정. 기본값은 모두 false.</summary>
        private void SetViewState(
            bool showResult = false, bool showNotAnalyzed = false, bool showAnalyzing = false)
        {
            _pnlPriorityBadge.Visible  = showResult;
            _lblSummaryHeader.Visible  = showResult;
            _lblSummary.Visible        = showResult;
            _lblReasonHeader.Visible   = showResult;
            _lblReason.Visible         = showResult;
            _pnlDivider.Visible        = showResult;
            _lblChangePriority.Visible = showResult;
            _cmbPriority.Visible       = showResult;
            _btnReanalyze.Visible      = showResult;
            _btnMoveFolder.Visible     = showResult;
            _pnlNotAnalyzed.Visible    = showNotAnalyzed;
            _pnlAnalyzing.Visible      = showAnalyzing;
        }

        private void RenderAnalysis(MailAnalysis analysis)
        {
            SetViewState(showResult: true);

            _pnlPriorityBadge.BackColor = GetPriorityColor(analysis.Priority);
            _lblPriorityText.Text = analysis.Priority.ToEmoji() + " " + analysis.Priority.ToKorean();

            _lblSummary.Text = analysis.IsFallback
                ? "(분석 실패) " + analysis.ErrorMessage
                : analysis.Summary ?? "";

            _lblReason.Text = analysis.PriorityReason ?? "";

            _cmbPriority.SelectedIndexChanged -= OnPriorityChanged;
            _cmbPriority.SelectedIndex = (int)analysis.Priority - 1;
            _cmbPriority.SelectedIndexChanged += OnPriorityChanged;
        }

        // ──────────────────────────────────────────────────────────────
        // 이벤트 핸들러
        // ──────────────────────────────────────────────────────────────

        private async void OnReanalyzeClick(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_currentEntryId)) return;

            Outlook.MailItem mail = null;
            try
            {
                mail = Globals.ThisAddIn.Application.Session.GetItemFromID(_currentEntryId)
                       as Outlook.MailItem;
                if (mail == null) return;

                ShowAnalyzing();

                // LLM_Analyzed 플래그 초기화 → 강제 재분석
                Services.MailProcessor.ClearAnalysisFlag(mail);

                var analysis = await Globals.ThisAddIn.MailProcessor.AnalyzeSingleAsync(mail);
                DisplayAnalysis(mail, analysis);
            }
            catch (Exception ex)
            {
                Logger.Error("OnReanalyzeClick: reanalysis failed", ex);
                ShowError(ex.Message);
            }
            finally
            {
                ComHelper.Release(mail);
            }
        }

        private void OnMoveFolderClick(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_currentEntryId)) return;

            Outlook.MailItem mail = null;
            try
            {
                mail = Globals.ThisAddIn.Application.Session.GetItemFromID(_currentEntryId)
                       as Outlook.MailItem;
                if (mail == null) return;

                var existing = Globals.ThisAddIn.MailProcessor.LoadExistingAnalysis(mail);
                if (existing == null)
                {
                    MessageBox.Show("분석 결과가 없습니다. 먼저 요약을 실행해 주세요.",
                        "MailPrioritizer", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                Globals.ThisAddIn.FolderManager.MoveToFolder(mail, existing.Priority);
            }
            finally
            {
                ComHelper.Release(mail);
            }
        }

        private void OnPriorityChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_currentEntryId)) return;

            Priority newPriority = (Priority)(_cmbPriority.SelectedIndex + 1);

            Outlook.MailItem mail = null;
            try
            {
                mail = Globals.ThisAddIn.Application.Session.GetItemFromID(_currentEntryId)
                       as Outlook.MailItem;
                if (mail == null) return;

                Outlook.UserProperties props = null;
                Outlook.UserProperty prop = null;
                try
                {
                    props = mail.UserProperties;
                    prop = props.Find(MailPropertyNames.Priority) ??
                           props.Add(MailPropertyNames.Priority, Outlook.OlUserPropertyType.olText);
                    prop.Value = newPriority.ToString();
                    mail.Save();
                }
                finally
                {
                    ComHelper.ReleaseAll(prop, props);
                }

                if (Globals.ThisAddIn.Config.Classification.AutoMoveToFolder)
                    Globals.ThisAddIn.FolderManager.MoveToFolder(mail, newPriority);

                _pnlPriorityBadge.BackColor = GetPriorityColor(newPriority);
                _lblPriorityText.Text = newPriority.ToEmoji() + " " + newPriority.ToKorean();
            }
            finally
            {
                ComHelper.Release(mail);
            }
        }

        private async void OnAnalyzeNowClick(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_currentEntryId)) return;

            Outlook.MailItem mail = null;
            try
            {
                mail = Globals.ThisAddIn.Application.Session.GetItemFromID(_currentEntryId)
                       as Outlook.MailItem;
                if (mail == null) return;

                ShowAnalyzing();
                var analysis = await Globals.ThisAddIn.MailProcessor.AnalyzeSingleAsync(mail);
                DisplayAnalysis(mail, analysis);
            }
            catch (Exception ex)
            {
                Logger.Error("OnAnalyzeNowClick: analysis failed", ex);
                ShowError(ex.Message);
            }
            finally
            {
                ComHelper.Release(mail);
            }
        }

        // ──────────────────────────────────────────────────────────────
        // 헬퍼
        // ──────────────────────────────────────────────────────────────

        private static Color GetPriorityColor(Priority priority)
        {
            switch (priority)
            {
                case Priority.Urgent: return Color.FromArgb(255, 235, 235);
                case Priority.High:   return Color.FromArgb(255, 243, 224);
                case Priority.Normal: return Color.FromArgb(232, 245, 233);
                case Priority.Low:    return Color.FromArgb(245, 245, 245);
                default:              return Color.White;
            }
        }

        /// <summary>받은편지함 통계를 갱신한다. COM 작업은 STA에서, UI 블로킹 최소화.</summary>
        public void RefreshStats()
        {
            _lblStatsContent.Text = "통계 로드 중...";
            _btnRefreshStats.Enabled = false;

            try
            {
                // STA 스레드에서 COM 작업 수행 (Outlook 이벤트 핸들러이므로 이미 STA)
                var stats = Globals.ThisAddIn.MailProcessor.CollectInboxStats(
                    Globals.ThisAddIn.Application);
                _lblStatsContent.Text = string.Format(
                    "전체: {0}건  분석됨: {1}건\n긴급: {2}  높음: {3}  보통: {4}  낮음: {5}",
                    stats.Total, stats.Analyzed,
                    stats.Urgent, stats.High, stats.Normal, stats.Low);
            }
            catch (Exception ex)
            {
                Logger.Error("RefreshStats: failed", ex);
                _lblStatsContent.Text = "(통계 로드 실패)";
            }
            finally
            {
                _btnRefreshStats.Enabled = true;
            }
        }

        private void SafeInvoke(Action action)
        {
            if (IsHandleCreated && InvokeRequired)
                Invoke(action);
            else
                action();
        }

        // ──────────────────────────────────────────────────────────────
        // UI 초기화
        // ──────────────────────────────────────────────────────────────

        private void InitializeComponent()
        {
            this.Dock = DockStyle.Fill;
            this.Font = new Font("맑은 고딕", 9f);
            this.AutoScroll = true;
            this.Padding = new Padding(8);

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                AutoSize = true
            };

            // 우선순위 배지
            _pnlPriorityBadge = new Panel { Height = 36, Dock = DockStyle.Fill, BackColor = Color.White };
            _lblPriorityText = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("맑은 고딕", 12f, FontStyle.Bold)
            };
            _pnlPriorityBadge.Controls.Add(_lblPriorityText);

            // 요약
            _lblSummaryHeader = new Label { Text = "요약:", Font = new Font("맑은 고딕", 9f, FontStyle.Bold), AutoSize = true };
            _lblSummary = new Label { Dock = DockStyle.Fill, AutoSize = false, Height = 80, Text = "", BorderStyle = BorderStyle.FixedSingle };

            // 판단 근거
            _lblReasonHeader = new Label { Text = "판단 근거:", Font = new Font("맑은 고딕", 9f, FontStyle.Bold), AutoSize = true };
            _lblReason = new Label { Dock = DockStyle.Fill, AutoSize = false, Height = 44, Text = "" };

            // 구분선
            _pnlDivider = new Panel { Height = 1, Dock = DockStyle.Fill, BackColor = Color.LightGray };

            // 우선순위 변경
            _lblChangePriority = new Label { Text = "우선순위 변경:", AutoSize = true };
            _cmbPriority = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 100 };
            _cmbPriority.Items.AddRange(new object[] { "긴급", "높음", "보통", "낮음" });

            // 버튼
            _btnReanalyze = new Button { Text = "재분석", Width = 80, Height = 28 };
            _btnMoveFolder = new Button { Text = "폴더이동", Width = 80, Height = 28 };
            _btnReanalyze.Click += OnReanalyzeClick;
            _btnMoveFolder.Click += OnMoveFolderClick;
            _cmbPriority.SelectedIndexChanged += OnPriorityChanged;

            var btnPanel = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            btnPanel.Controls.AddRange(new Control[] { _btnReanalyze, _btnMoveFolder });

            // 미분석 상태 패널
            _pnlNotAnalyzed = new Panel { Dock = DockStyle.Fill, Height = 80 };
            _lblNotAnalyzed = new Label
            {
                Dock = DockStyle.Fill,
                ForeColor = Color.Gray,
                TextAlign = ContentAlignment.MiddleCenter
            };
            _btnAnalyzeNow = new Button { Text = "지금 분석하기", Dock = DockStyle.Bottom, Height = 28 };
            _btnAnalyzeNow.Click += OnAnalyzeNowClick;
            _pnlNotAnalyzed.Controls.AddRange(new Control[] { _lblNotAnalyzed, _btnAnalyzeNow });

            // 분석 중 패널
            _pnlAnalyzing = new Panel { Dock = DockStyle.Fill, Height = 40 };
            _lblAnalyzing = new Label
            {
                Text = "분석 중...",
                Dock = DockStyle.Fill,
                ForeColor = Color.CornflowerBlue,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("맑은 고딕", 10f)
            };
            _pnlAnalyzing.Controls.Add(_lblAnalyzing);

            layout.Controls.Add(_pnlPriorityBadge);
            layout.Controls.Add(_lblSummaryHeader);
            layout.Controls.Add(_lblSummary);
            layout.Controls.Add(_lblReasonHeader);
            layout.Controls.Add(_lblReason);
            layout.Controls.Add(_pnlDivider);
            layout.Controls.Add(_lblChangePriority);
            layout.Controls.Add(_cmbPriority);
            layout.Controls.Add(btnPanel);
            layout.Controls.Add(_pnlNotAnalyzed);
            layout.Controls.Add(_pnlAnalyzing);

            // 통계 패널 (#9)
            var statsDivider = new Panel { Height = 1, Dock = DockStyle.Fill, BackColor = Color.LightGray };
            _pnlStats = new Panel { Dock = DockStyle.Fill, Height = 80 };
            _lblStatsTitle = new Label
            {
                Text = "받은편지함 통계",
                Font = new Font("맑은 고딕", 9f, FontStyle.Bold),
                AutoSize = true,
                Dock = DockStyle.Top
            };
            _lblStatsContent = new Label
            {
                Text = "(새로고침 버튼을 눌러 통계를 확인하세요)",
                Dock = DockStyle.Fill,
                ForeColor = Color.DimGray,
                AutoSize = false,
                Height = 40
            };
            _btnRefreshStats = new Button { Text = "통계 새로고침", Width = 100, Height = 24, Dock = DockStyle.Bottom };
            _btnRefreshStats.Click += (s, ev) => RefreshStats();
            _pnlStats.Controls.Add(_lblStatsContent);
            _pnlStats.Controls.Add(_lblStatsTitle);
            _pnlStats.Controls.Add(_btnRefreshStats);

            layout.Controls.Add(statsDivider);
            layout.Controls.Add(_pnlStats);

            this.Controls.Add(layout);
        }
    }
}
