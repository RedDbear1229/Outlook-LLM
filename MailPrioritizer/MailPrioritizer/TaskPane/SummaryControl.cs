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
        // ── 컨트롤 필드 ─────────────────────────────────────────────────
        private Panel           _pnlPriorityBadge;
        private Label           _lblPriorityText;
        private Label           _lblSummaryHeader;
        private Label           _lblSummary;
        private Label           _lblReasonHeader;
        private Label           _lblReason;
        private Panel           _pnlDivider;
        private Label           _lblChangePriority;
        private ComboBox        _cmbPriority;
        private Button          _btnReanalyze;
        private Button          _btnMoveFolder;
        private Panel           _pnlNotAnalyzed;
        private Label           _lblNotAnalyzed;
        private Button          _btnAnalyzeNow;
        private Panel           _pnlAnalyzing;
        private Label           _lblAnalyzing;
        private Panel           _pnlStats;
        private Label           _lblStatsTitle;
        private Label           _lblStatsContent;
        private Button          _btnRefreshStats;
        private ToolTip         _toolTip;
        private Label           _lblApiStatus;

        // ApplyTheme 에서 접근하기 위해 필드로 보관
        private TableLayoutPanel _layout;
        private FlowLayoutPanel  _btnPanel;
        private Panel            _statsDivider;

        // ── 상태 ────────────────────────────────────────────────────────
        private string       _currentEntryId;
        private MailAnalysis _currentAnalysis;
        private ThemePalette _currentTheme = ThemePalette.Light;

        // ── 생성자 ──────────────────────────────────────────────────────
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
                _currentEntryId  = mail?.EntryID;
                _currentAnalysis = analysis;
                RenderAnalysis(analysis);
            });
        }

        public void ShowNoSelection()
        {
            SafeInvoke(() =>
            {
                SetViewState(showNotAnalyzed: true);
                _lblNotAnalyzed.Text    = "메일을 선택하면 분석 결과가 표시됩니다.";
                _btnAnalyzeNow.Visible  = false;
            });
        }

        public void ShowUnanalyzed(Outlook.MailItem mail)
        {
            SafeInvoke(() =>
            {
                _currentEntryId       = mail?.EntryID;
                SetViewState(showNotAnalyzed: true);
                _lblNotAnalyzed.Text   = "아직 분석되지 않은 메일입니다.";
                _btnAnalyzeNow.Visible = true;
            });
        }

        public void ShowError(string message)
        {
            SafeInvoke(() =>
            {
                SetViewState(showNotAnalyzed: true);
                _lblNotAnalyzed.Text   = "오류: " + message;
                _btnAnalyzeNow.Visible = false;
            });
        }

        /// <summary>R-08: 키보드 단축키로 우선순위 변경 시 배지만 즉시 갱신.</summary>
        public void RefreshPriorityBadge(Priority priority)
        {
            SafeInvoke(() =>
            {
                if (!_pnlPriorityBadge.Visible) return;
                _pnlPriorityBadge.BackColor = GetBadgeBack(priority);
                _lblPriorityText.ForeColor  = GetBadgeText(priority);
                _lblPriorityText.Text       = priority.ToEmoji() + " " + priority.ToKorean();
            });
        }

        /// <summary>R-13: LLM 서비스 헬스 상태를 하단 레이블에 갱신.</summary>
        public void UpdateApiStatus()
        {
            SafeInvoke(() =>
            {
                var svc = Globals.ThisAddIn.LlmService;
                if (svc == null) return;

                if (svc.LastErrorAt.HasValue
                    && (!svc.LastSuccessAt.HasValue || svc.LastErrorAt > svc.LastSuccessAt))
                {
                    string when = svc.LastErrorAt.Value.ToString("HH:mm");
                    string msg  = svc.LastErrorMessage ?? "";
                    if (msg.Length > 50) msg = msg.Substring(0, 50) + "...";
                    _lblApiStatus.Text      = "API 오류 " + when + ": " + msg;
                    _lblApiStatus.ForeColor = Color.FromArgb(180, 0, 0);
                }
                else if (svc.LastSuccessAt.HasValue)
                {
                    string when = svc.LastSuccessAt.Value.ToString("HH:mm");
                    _lblApiStatus.Text      = "마지막 성공: " + when;
                    _lblApiStatus.ForeColor = Color.FromArgb(0, 130, 0);
                }
                else
                {
                    _lblApiStatus.Text      = "API 미호출";
                    _lblApiStatus.ForeColor = _currentTheme.HintText;
                }
            });
        }

        // ──────────────────────────────────────────────────────────────
        // 테마 적용
        // ──────────────────────────────────────────────────────────────

        /// <summary>
        /// Task Pane 전체에 지정된 테마 팔레트를 적용합니다.
        /// ThisAddIn.ApplyConfigChange() 또는 Startup에서 호출합니다.
        /// </summary>
        internal void ApplyTheme(ThemePalette palette)
        {
            SafeInvoke(() =>
            {
                _currentTheme = palette;

                // 루트 및 레이아웃 배경
                BackColor        = palette.PanelBackground;
                _layout.BackColor = palette.PanelBackground;

                // 우선순위 배지
                _pnlPriorityBadge.BackColor = palette.PanelBackground;
                if (_currentAnalysis != null)
                {
                    _pnlPriorityBadge.BackColor = GetBadgeBack(_currentAnalysis.Priority);
                    _lblPriorityText.ForeColor   = GetBadgeText(_currentAnalysis.Priority);
                }
                else
                {
                    _lblPriorityText.ForeColor = palette.PrimaryText;
                }

                // 요약 영역
                _lblSummaryHeader.ForeColor = palette.PrimaryText;
                _lblSummaryHeader.BackColor = palette.PanelBackground;
                _lblSummary.ForeColor       = palette.PrimaryText;
                _lblSummary.BackColor       = palette.ControlBackground;

                // 판단 근거
                _lblReasonHeader.ForeColor = palette.PrimaryText;
                _lblReasonHeader.BackColor = palette.PanelBackground;
                _lblReason.ForeColor       = palette.SecondaryText;
                _lblReason.BackColor       = palette.PanelBackground;

                // 구분선
                _pnlDivider.BackColor   = palette.Divider;
                _statsDivider.BackColor = palette.Divider;

                // 우선순위 변경 컨트롤
                _lblChangePriority.ForeColor = palette.PrimaryText;
                _lblChangePriority.BackColor = palette.PanelBackground;
                _cmbPriority.BackColor       = palette.ComboBack;
                _cmbPriority.ForeColor       = palette.ComboText;

                // 버튼
                ApplyButtonTheme(_btnReanalyze,    palette);
                ApplyButtonTheme(_btnMoveFolder,   palette);
                ApplyButtonTheme(_btnAnalyzeNow,   palette);
                ApplyButtonTheme(_btnRefreshStats, palette);
                _btnPanel.BackColor = palette.PanelBackground;

                // 미분석 패널
                _pnlNotAnalyzed.BackColor   = palette.PanelBackground;
                _lblNotAnalyzed.ForeColor   = palette.HintText;
                _lblNotAnalyzed.BackColor   = palette.PanelBackground;

                // 분석 중 패널
                _pnlAnalyzing.BackColor   = palette.PanelBackground;
                _lblAnalyzing.ForeColor   = palette.AnalyzingText;
                _lblAnalyzing.BackColor   = palette.PanelBackground;

                // 통계 패널
                _pnlStats.BackColor       = palette.PanelBackground;
                _lblStatsTitle.ForeColor  = palette.PrimaryText;
                _lblStatsTitle.BackColor  = palette.PanelBackground;
                _lblStatsContent.ForeColor = palette.SecondaryText;
                _lblStatsContent.BackColor = palette.PanelBackground;

                // API 상태 레이블
                _lblApiStatus.BackColor = palette.PanelBackground;

                Invalidate(true);
            });
        }

        private static void ApplyButtonTheme(Button btn, ThemePalette palette)
        {
            btn.BackColor = palette.ButtonBackground;
            btn.ForeColor = palette.ButtonText;
            btn.UseVisualStyleBackColor = false;
            if (palette.ButtonFlat)
            {
                btn.FlatStyle = FlatStyle.Flat;
                btn.FlatAppearance.BorderColor = palette.Border;
            }
            else
            {
                btn.FlatStyle = FlatStyle.Standard;
            }
        }

        // ──────────────────────────────────────────────────────────────
        // 렌더링
        // ──────────────────────────────────────────────────────────────

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

            _pnlPriorityBadge.BackColor = GetBadgeBack(analysis.Priority);
            _lblPriorityText.ForeColor  = GetBadgeText(analysis.Priority);
            _lblPriorityText.Text = analysis.Priority.ToEmoji() + " " + analysis.Priority.ToKorean();

            if (analysis.IsFallback)
                _lblSummary.Text = "(분석 실패) " + analysis.ErrorMessage;
            else if (analysis.IsRuleBased)
                _lblSummary.Text = "(발신자 규칙 적용)";
            else
                _lblSummary.Text = analysis.Summary ?? "";

            _lblReason.Text = analysis.PriorityReason ?? "";

            // R-07: 분석 모델/출처 정보 (툴팁)
            string tooltip = analysis.IsRuleBased ? "규칙 기반 분류"
                : !string.IsNullOrEmpty(analysis.ModelName)
                    ? "모델: " + analysis.ModelName
                      + "  분석일: " + analysis.AnalyzedAt.ToString("MM-dd HH:mm")
                    : "";
            _toolTip.SetToolTip(_lblSummaryHeader, tooltip);

            _cmbPriority.SelectedIndexChanged -= OnPriorityChanged;
            _cmbPriority.SelectedIndex = (int)analysis.Priority - 1;
            _cmbPriority.SelectedIndexChanged += OnPriorityChanged;
        }

        // ──────────────────────────────────────────────────────────────
        // 이벤트 핸들러
        // ──────────────────────────────────────────────────────────────

        private async void OnReanalyzeClick(object sender, EventArgs e)
        {
            await RunAnalysisAsync(clearFirst: true);
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
                Outlook.UserProperty prop    = null;
                try
                {
                    props = mail.UserProperties;
                    prop  = props.Find(MailPropertyNames.Priority) ??
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

                // R-06: 피드백 기록
                if (_currentAnalysis != null && !_currentAnalysis.IsRuleBased
                    && !_currentAnalysis.IsFallback
                    && _currentAnalysis.Priority != newPriority)
                {
                    Globals.ThisAddIn.FeedbackStore.Record(
                        _currentEntryId,
                        mail.Subject ?? "",
                        _currentAnalysis.Priority,
                        newPriority);
                    _currentAnalysis = null;
                }

                _pnlPriorityBadge.BackColor = GetBadgeBack(newPriority);
                _lblPriorityText.ForeColor  = GetBadgeText(newPriority);
                _lblPriorityText.Text       = newPriority.ToEmoji() + " " + newPriority.ToKorean();
            }
            finally
            {
                ComHelper.Release(mail);
            }
        }

        private async void OnAnalyzeNowClick(object sender, EventArgs e)
        {
            await RunAnalysisAsync(clearFirst: false);
        }

        private async Task RunAnalysisAsync(bool clearFirst)
        {
            if (string.IsNullOrEmpty(_currentEntryId)) return;

            Outlook.MailItem mail = null;
            try
            {
                mail = Globals.ThisAddIn.Application.Session.GetItemFromID(_currentEntryId)
                       as Outlook.MailItem;
                if (mail == null) return;

                if (clearFirst)
                    Services.MailProcessor.ClearAnalysisFlag(mail);

                ShowAnalyzing();
                var analysis = await Globals.ThisAddIn.MailProcessor.AnalyzeSingleAsync(mail);
                DisplayAnalysis(mail, analysis);
                UpdateApiStatus();
            }
            catch (Exception ex)
            {
                Logger.Error("RunAnalysisAsync: analysis failed", ex);
                ShowError(ex.Message);
                UpdateApiStatus();
            }
            finally
            {
                ComHelper.Release(mail);
            }
        }

        // ──────────────────────────────────────────────────────────────
        // 헬퍼
        // ──────────────────────────────────────────────────────────────

        private Color GetBadgeBack(Priority priority)
        {
            switch (priority)
            {
                case Priority.Urgent: return _currentTheme.UrgentBadge;
                case Priority.High:   return _currentTheme.HighBadge;
                case Priority.Normal: return _currentTheme.NormalBadge;
                case Priority.Low:    return _currentTheme.LowBadge;
                default:              return _currentTheme.PanelBackground;
            }
        }

        private Color GetBadgeText(Priority priority)
        {
            switch (priority)
            {
                case Priority.Urgent: return _currentTheme.UrgentBadgeText;
                case Priority.High:   return _currentTheme.HighBadgeText;
                case Priority.Normal: return _currentTheme.NormalBadgeText;
                case Priority.Low:    return _currentTheme.LowBadgeText;
                default:              return _currentTheme.PrimaryText;
            }
        }

        /// <summary>받은편지함 통계를 갱신한다.</summary>
        public void RefreshStats()
        {
            _lblStatsContent.Text    = "통계 로드 중...";
            _btnRefreshStats.Enabled = false;

            try
            {
                var stats    = Globals.ThisAddIn.MailProcessor.CollectInboxStats(Globals.ThisAddIn.Application);
                var fbStats  = Globals.ThisAddIn.FeedbackStore.GetStats(stats.Analyzed);

                string accuracyLine = stats.Analyzed > 0
                    ? string.Format("LLM 정확도: {0}%  수동 변경: {1}건",
                        fbStats.AccuracyPercent, fbStats.CorrectionCount)
                    : "";
                string misclassifiedLine = !string.IsNullOrEmpty(fbStats.MostMisclassified)
                    ? "최다 오분류: " + fbStats.MostMisclassified
                    : "";

                _lblStatsContent.Text = string.Format(
                    "전체: {0}건  분석됨: {1}건\n긴급: {2}  높음: {3}  보통: {4}  낮음: {5}"
                    + (accuracyLine != "" ? "\n" + accuracyLine : "")
                    + (misclassifiedLine != "" ? "\n" + misclassifiedLine : ""),
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
            if (IsDisposed || !IsHandleCreated) return;
            if (InvokeRequired)
            {
                try { Invoke(action); }
                catch (ObjectDisposedException) { }
                catch (InvalidOperationException) { }
            }
            else
            {
                action();
            }
        }

        // ──────────────────────────────────────────────────────────────
        // UI 초기화
        // ──────────────────────────────────────────────────────────────

        private void InitializeComponent()
        {
            Dock      = DockStyle.Fill;
            Font      = new Font("맑은 고딕", 9f);
            AutoScroll = true;
            Padding   = new Padding(8);

            _toolTip = new ToolTip { AutoPopDelay = 8000, InitialDelay = 500 };

            _layout = new TableLayoutPanel
            {
                Dock        = DockStyle.Fill,
                ColumnCount = 1,
                AutoSize    = true
            };

            // 우선순위 배지
            _pnlPriorityBadge = new Panel { Height = 36, Dock = DockStyle.Fill };
            _lblPriorityText  = new Label
            {
                Dock      = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font      = new Font("맑은 고딕", 12f, FontStyle.Bold)
            };
            _pnlPriorityBadge.Controls.Add(_lblPriorityText);

            // 요약
            _lblSummaryHeader = new Label { Text = "요약:", Font = new Font("맑은 고딕", 9f, FontStyle.Bold), AutoSize = true };
            _lblSummary       = new Label { Dock = DockStyle.Fill, AutoSize = false, Height = 80, Text = "", BorderStyle = BorderStyle.FixedSingle };

            // 판단 근거
            _lblReasonHeader = new Label { Text = "판단 근거:", Font = new Font("맑은 고딕", 9f, FontStyle.Bold), AutoSize = true };
            _lblReason       = new Label { Dock = DockStyle.Fill, AutoSize = false, Height = 44, Text = "" };

            // 구분선
            _pnlDivider = new Panel { Height = 1, Dock = DockStyle.Fill, BackColor = Color.LightGray };

            // 우선순위 변경
            _lblChangePriority = new Label { Text = "우선순위 변경:", AutoSize = true };
            _cmbPriority       = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 100 };
            _cmbPriority.Items.AddRange(new object[] { "긴급", "높음", "보통", "낮음" });

            // 버튼
            _btnReanalyze = new Button { Text = "재분석",   Width = 80, Height = 28 };
            _btnMoveFolder = new Button { Text = "폴더이동", Width = 80, Height = 28 };
            _btnReanalyze.Click  += OnReanalyzeClick;
            _btnMoveFolder.Click += OnMoveFolderClick;
            _cmbPriority.SelectedIndexChanged += OnPriorityChanged;

            _btnPanel = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            _btnPanel.Controls.AddRange(new Control[] { _btnReanalyze, _btnMoveFolder });

            // 미분석 패널
            _pnlNotAnalyzed = new Panel { Dock = DockStyle.Fill, Height = 80 };
            _lblNotAnalyzed = new Label
            {
                Dock      = DockStyle.Fill,
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
                Text      = "분석 중...",
                Dock      = DockStyle.Fill,
                ForeColor = Color.CornflowerBlue,
                TextAlign = ContentAlignment.MiddleCenter,
                Font      = new Font("맑은 고딕", 10f)
            };
            _pnlAnalyzing.Controls.Add(_lblAnalyzing);

            _layout.Controls.Add(_pnlPriorityBadge);
            _layout.Controls.Add(_lblSummaryHeader);
            _layout.Controls.Add(_lblSummary);
            _layout.Controls.Add(_lblReasonHeader);
            _layout.Controls.Add(_lblReason);
            _layout.Controls.Add(_pnlDivider);
            _layout.Controls.Add(_lblChangePriority);
            _layout.Controls.Add(_cmbPriority);
            _layout.Controls.Add(_btnPanel);
            _layout.Controls.Add(_pnlNotAnalyzed);
            _layout.Controls.Add(_pnlAnalyzing);

            // 통계 패널
            _statsDivider = new Panel { Height = 1, Dock = DockStyle.Fill, BackColor = Color.LightGray };
            _pnlStats     = new Panel { Dock = DockStyle.Fill, Height = 80 };
            _lblStatsTitle = new Label
            {
                Text     = "받은편지함 통계",
                Font     = new Font("맑은 고딕", 9f, FontStyle.Bold),
                AutoSize = true,
                Dock     = DockStyle.Top
            };
            _lblStatsContent = new Label
            {
                Text      = "(새로고침 버튼을 눌러 통계를 확인하세요)",
                Dock      = DockStyle.Fill,
                ForeColor = Color.DimGray,
                AutoSize  = false,
                Height    = 40
            };
            _btnRefreshStats = new Button { Text = "통계 새로고침", Width = 100, Height = 24, Dock = DockStyle.Bottom };
            _btnRefreshStats.Click += (s, ev) => RefreshStats();
            _pnlStats.Controls.Add(_lblStatsContent);
            _pnlStats.Controls.Add(_lblStatsTitle);
            _pnlStats.Controls.Add(_btnRefreshStats);

            _layout.Controls.Add(_statsDivider);
            _layout.Controls.Add(_pnlStats);

            // R-13: API 상태 레이블 (하단)
            _lblApiStatus = new Label
            {
                Dock      = DockStyle.Fill,
                AutoSize  = false,
                Height    = 18,
                ForeColor = Color.Gray,
                Text      = "API 미호출",
                Font      = new Font("맑은 고딕", 8f),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding   = new Padding(2, 0, 0, 0)
            };
            _layout.Controls.Add(_lblApiStatus);

            Controls.Add(_layout);
        }
    }
}
