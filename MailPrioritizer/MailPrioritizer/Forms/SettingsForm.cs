using System;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using MailPrioritizer.Models;
using MailPrioritizer.Services;
using MailPrioritizer.Utils;

namespace MailPrioritizer.Forms
{
    /// <summary>환경설정 다이얼로그 (3탭: API 설정 / 프롬프트 / 분류 설정).</summary>
    public class SettingsForm : Form
    {
        // ── 반환값 ──
        public AppConfig ResultConfig { get; private set; }
        public string PlainToken { get; private set; }

        // ── 탭1: API 설정 ──
        private TextBox txtEndpoint;
        private TextBox txtToken;
        private TextBox txtModel;
        private Button btnTest;
        private Label lblTestResult;

        // ── 탭2: 프롬프트 ──
        private TextBox txtPrompt;
        private Button btnResetPrompt;

        // ── 탭3: 분류 설정 ──
        private TextBox txtFolderPrefix;
        private CheckBox chkAutoMove;
        private CheckBox chkTagSubject;
        private NumericUpDown numMaxBody;
        private ComboBox cmbStore;
        private CheckBox chkAutoAnalyze;
        private CheckBox chkIncludeAttachments;

        // ── 탭4: 발신자 규칙 ──
        private DataGridView dgvRules;

        private readonly AppConfig _original;

        // 단일 정의 출처: LlmService.DefaultSystemPrompt
        private static string DefaultPrompt => Services.LlmService.DefaultSystemPrompt;

        public SettingsForm(AppConfig current)
        {
            _original = current;
            InitializeComponent();
            LoadConfig(current);
        }

        private void InitializeComponent()
        {
            this.Text = "환경설정";
            this.Size = new Size(560, 580);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Font = new Font("맑은 고딕", 9f);

            var tabControl = new TabControl { Dock = DockStyle.Fill };
            var tab1 = new TabPage("API 설정");
            var tab2 = new TabPage("프롬프트");
            var tab3 = new TabPage("분류 설정");
            var tab4 = new TabPage("발신자 규칙");
            tabControl.TabPages.AddRange(new[] { tab1, tab2, tab3, tab4 });

            // ── 탭1: API 설정 ──
            tab1.Padding = new Padding(12);
            var pnl1 = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 6 };
            pnl1.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
            pnl1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            txtEndpoint = new TextBox { Dock = DockStyle.Fill };
            txtToken    = new TextBox { Dock = DockStyle.Fill, UseSystemPasswordChar = false };
            txtModel    = new TextBox { Dock = DockStyle.Fill };
            btnTest     = new Button { Text = "연결 테스트", Width = 110, Height = 30 };
            lblTestResult = new Label { Dock = DockStyle.Fill, ForeColor = Color.Gray, TextAlign = ContentAlignment.MiddleLeft };

            btnTest.Click += OnTestConnectionClick;

            AddRow(pnl1, "Endpoint URL:", txtEndpoint);
            AddRow(pnl1, "API Token:", txtToken);
            AddRow(pnl1, "모델 이름:", txtModel);

            var testPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight };
            testPanel.Controls.Add(btnTest);
            testPanel.Controls.Add(lblTestResult);
            pnl1.Controls.Add(new Label { Text = "", Dock = DockStyle.Fill });
            pnl1.Controls.Add(testPanel);

            var hintLabel = new Label
            {
                Text = "예: https://api.anthropic.com/v1  또는  https://api.openai.com/v1",
                ForeColor = Color.Gray,
                Dock = DockStyle.Fill,
                AutoSize = false
            };
            pnl1.SetColumnSpan(hintLabel, 2);
            pnl1.Controls.Add(hintLabel);

            tab1.Controls.Add(pnl1);

            // ── 탭2: 프롬프트 ──
            tab2.Padding = new Padding(12);
            var pnl2 = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3 };
            pnl2.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            pnl2.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            pnl2.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            txtPrompt = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Consolas", 9f)
            };

            var hintPrompt = new Label
            {
                Text = "※ JSON 응답 형식 부분({\"summary\":...,\"priority\":...,\"priority_reason\":...})은 반드시 유지하세요.",
                ForeColor = Color.DarkRed,
                AutoSize = false,
                Dock = DockStyle.Fill,
                Height = 36
            };

            btnResetPrompt = new Button { Text = "기본값으로 복원", AutoSize = true };
            btnResetPrompt.Click += (s, e) => { txtPrompt.Text = DefaultPrompt; };

            pnl2.Controls.Add(hintPrompt);
            pnl2.Controls.Add(txtPrompt);
            pnl2.Controls.Add(btnResetPrompt);
            tab2.Controls.Add(pnl2);

            // ── 탭3: 분류 설정 ──
            tab3.Padding = new Padding(12);
            var pnl3 = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 6 };
            pnl3.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140));
            pnl3.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            txtFolderPrefix = new TextBox { Dock = DockStyle.Fill };
            chkAutoMove     = new CheckBox { Text = "분석 후 자동으로 폴더 이동", Dock = DockStyle.Fill, AutoSize = false };
            chkTagSubject   = new CheckBox { Text = "제목에 우선순위 태그 삽입 ([긴급] 등)", Dock = DockStyle.Fill, AutoSize = false };
            numMaxBody      = new NumericUpDown { Minimum = 500, Maximum = 20000, Increment = 500, Dock = DockStyle.Fill };

            var folderHint = new Label
            {
                Text = "※ 일부 서버는 이모지 폴더명 미지원. 한글/영문 권장.",
                ForeColor = Color.Gray,
                Dock = DockStyle.Fill,
                AutoSize = false
            };
            pnl3.SetColumnSpan(folderHint, 2);

            AddRow(pnl3, "폴더 접두사:", txtFolderPrefix);
            pnl3.Controls.Add(folderHint);
            pnl3.Controls.Add(new Label());
            pnl3.Controls.Add(chkAutoMove);
            pnl3.Controls.Add(new Label());
            pnl3.Controls.Add(chkTagSubject);
            AddRow(pnl3, "본문 최대 길이:", numMaxBody);

            // 대상 저장소 (#10)
            cmbStore = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            AddRow(pnl3, "대상 계정:", cmbStore);
            LoadStoreList();

            // 신규 메일 자동 분석 (#11)
            chkAutoAnalyze = new CheckBox { Text = "새 메일 수신 시 자동 분석", Dock = DockStyle.Fill, AutoSize = false };
            pnl3.Controls.Add(new Label());
            pnl3.Controls.Add(chkAutoAnalyze);

            // 첨부파일명 포함 (R-04)
            chkIncludeAttachments = new CheckBox { Text = "첨부파일명을 분석 프롬프트에 포함", Dock = DockStyle.Fill, AutoSize = false };
            pnl3.Controls.Add(new Label());
            pnl3.Controls.Add(chkIncludeAttachments);

            tab3.Controls.Add(pnl3);

            // ── 탭4: 발신자 규칙 (R-01) ──
            tab4.Padding = new Padding(8);
            var pnl4 = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3 };
            pnl4.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            pnl4.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            pnl4.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var ruleHint = new Label
            {
                Text = "발신자 이메일 또는 도메인에 규칙을 설정하면 LLM 호출 없이 즉시 분류됩니다.",
                ForeColor = Color.DimGray,
                Dock = DockStyle.Fill,
                AutoSize = false,
                Height = 32
            };

            dgvRules = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = true,
                AllowUserToDeleteRows = true,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                Font = new Font("맑은 고딕", 8.5f)
            };

            var colEnabled = new DataGridViewCheckBoxColumn
            {
                Name = "Enabled", HeaderText = "활성", Width = 44, FillWeight = 15
            };
            var colType = new DataGridViewComboBoxColumn
            {
                Name = "Type", HeaderText = "유형", FillWeight = 20
            };
            colType.Items.AddRange(new object[] { "email", "domain" });
            var colPattern = new DataGridViewTextBoxColumn
            {
                Name = "Pattern", HeaderText = "패턴 (이메일 또는 도메인)", FillWeight = 35
            };
            var colPriority = new DataGridViewComboBoxColumn
            {
                Name = "Priority", HeaderText = "우선순위", FillWeight = 15
            };
            colPriority.Items.AddRange(new object[] { "urgent", "high", "normal", "low" });
            var colNote = new DataGridViewTextBoxColumn
            {
                Name = "Note", HeaderText = "메모", FillWeight = 15
            };
            dgvRules.Columns.AddRange(new DataGridViewColumn[]
                { colEnabled, colType, colPattern, colPriority, colNote });

            var btnRowPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight };
            var btnAddRule = new Button { Text = "+ 규칙 추가", Width = 90, Height = 26 };
            var btnRemoveRule = new Button { Text = "선택 삭제", Width = 80, Height = 26 };
            btnAddRule.Click += (s, ev) =>
            {
                int row = dgvRules.Rows.Add(true, "email", "", "normal", "");
                dgvRules.CurrentCell = dgvRules.Rows[row].Cells["Pattern"];
            };
            btnRemoveRule.Click += (s, ev) =>
            {
                foreach (DataGridViewRow row in dgvRules.SelectedRows)
                    if (!row.IsNewRow) dgvRules.Rows.Remove(row);
            };
            btnRowPanel.Controls.AddRange(new Control[] { btnAddRule, btnRemoveRule });

            pnl4.Controls.Add(ruleHint);
            pnl4.Controls.Add(dgvRules);
            pnl4.Controls.Add(btnRowPanel);
            tab4.Controls.Add(pnl4);

            // ── 하단 버튼 ──
            var btnPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.RightToLeft,
                Height = 44,
                Padding = new Padding(8)
            };
            var btnCancel = new Button { Text = "취소", DialogResult = DialogResult.Cancel, Width = 72 };
            var btnSave   = new Button { Text = "저장",  DialogResult = DialogResult.None,   Width = 72 };
            btnSave.Click += OnSaveClick;
            btnPanel.Controls.AddRange(new Control[] { btnCancel, btnSave });

            this.Controls.Add(tabControl);
            this.Controls.Add(btnPanel);
            this.CancelButton = btnCancel;
        }

        private static void AddRow(TableLayoutPanel panel, string labelText, Control control)
        {
            panel.Controls.Add(new Label
            {
                Text = labelText,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleRight,
                AutoSize = false
            });
            panel.Controls.Add(control);
        }

        private void LoadConfig(AppConfig config)
        {
            txtEndpoint.Text     = config.Llm.Endpoint;
            txtToken.Text        = config.Llm.ApiToken;
            txtModel.Text        = config.Llm.ModelName;
            txtPrompt.Text       = string.IsNullOrEmpty(config.Llm.SystemPrompt)
                                   ? DefaultPrompt : config.Llm.SystemPrompt;
            txtFolderPrefix.Text = config.Classification.FolderPrefix;
            chkAutoMove.Checked  = config.Classification.AutoMoveToFolder;
            chkTagSubject.Checked = config.Display.TagSubjectWithPriority;
            numMaxBody.Value     = config.Processing.MaxBodyLength;
            chkAutoAnalyze.Checked = config.Processing.AutoAnalyzeNewMail;
            chkIncludeAttachments.Checked = config.Processing.IncludeAttachmentNames;

            // 발신자 규칙 복원 (R-01)
            dgvRules.Rows.Clear();
            if (config.Rules != null)
            {
                foreach (var rule in config.Rules.SenderRules)
                    dgvRules.Rows.Add(rule.Enabled, rule.Type, rule.Pattern, rule.Priority, rule.Note);
            }

            // 저장소 선택 복원
            string targetId = config.Processing.TargetStoreId ?? "";
            for (int i = 0; i < cmbStore.Items.Count; i++)
            {
                var item = cmbStore.Items[i] as StoreItem;
                if (item != null && item.StoreId == targetId)
                {
                    cmbStore.SelectedIndex = i;
                    return;
                }
            }
            if (cmbStore.Items.Count > 0)
                cmbStore.SelectedIndex = 0;
        }

        private void LoadStoreList()
        {
            cmbStore.Items.Clear();
            cmbStore.Items.Add(new StoreItem("(기본 저장소)", ""));

            Microsoft.Office.Interop.Outlook.NameSpace session = null;
            Microsoft.Office.Interop.Outlook.Stores stores = null;
            try
            {
                session = Globals.ThisAddIn.Application.Session;
                stores = session.Stores;
                for (int i = 1; i <= stores.Count; i++)
                {
                    Microsoft.Office.Interop.Outlook.Store store = null;
                    try
                    {
                        store = stores[i];
                        cmbStore.Items.Add(new StoreItem(store.DisplayName, store.StoreID));
                    }
                    finally
                    {
                        ComHelper.Release(store);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warn("LoadStoreList: failed to enumerate stores - " + ex.Message);
            }
            finally
            {
                ComHelper.ReleaseAll(stores, session);
            }

            if (cmbStore.Items.Count > 0)
                cmbStore.SelectedIndex = 0;
        }

        /// <summary>ComboBox 항목용 저장소 래퍼.</summary>
        private class StoreItem
        {
            public string DisplayName { get; }
            public string StoreId { get; }
            public StoreItem(string displayName, string storeId)
            {
                DisplayName = displayName;
                StoreId = storeId;
            }
            public override string ToString() { return DisplayName; }
        }

        private async void OnTestConnectionClick(object sender, EventArgs e)
        {
            btnTest.Enabled = false;
            lblTestResult.Text = "테스트 중...";
            lblTestResult.ForeColor = Color.Gray;

            try
            {
                var testConfig = BuildConfig();
                testConfig.Llm.ApiToken = txtToken.Text;
                using (var svc = new LlmService(testConfig))
                {
                    var result = await svc.TestConnectionAsync();
                    lblTestResult.Text      = result.Item1 ? "✅ " + result.Item2 : "❌ " + result.Item2;
                    lblTestResult.ForeColor = result.Item1 ? Color.Green : Color.Red;
                }
            }
            catch (Exception ex)
            {
                Logger.Error("OnTestConnectionClick: test failed", ex);
                lblTestResult.Text      = "❌ 오류: " + ex.Message;
                lblTestResult.ForeColor = Color.Red;
            }
            finally
            {
                btnTest.Enabled = true;
            }
        }

        private void OnSaveClick(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtEndpoint.Text))
            {
                MessageBox.Show("Endpoint URL을 입력해 주세요.", "MailPrioritizer",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            PlainToken   = txtToken.Text;
            ResultConfig = BuildConfig();

            // 시스템 프롬프트: 기본값이면 빈 문자열로 저장
            if (ResultConfig.Llm.SystemPrompt.Trim() == DefaultPrompt.Trim())
                ResultConfig.Llm.SystemPrompt = "";

            Logger.Info("Settings saved by user");
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private AppConfig BuildConfig()
        {
            var config = new AppConfig();
            config.Llm.Endpoint     = txtEndpoint.Text.Trim();
            config.Llm.ModelName    = txtModel.Text.Trim();
            config.Llm.SystemPrompt = txtPrompt.Text;
            config.Llm.MaxTokens    = _original.Llm.MaxTokens;
            config.Llm.TimeoutSeconds = _original.Llm.TimeoutSeconds;

            config.Classification.FolderPrefix      = txtFolderPrefix.Text.Trim();
            config.Classification.AutoMoveToFolder   = chkAutoMove.Checked;
            config.Classification.FolderNames        = _original.Classification.FolderNames;

            config.Display.TagSubjectWithPriority = chkTagSubject.Checked;
            config.Display.TaskPaneVisible = _original.Display.TaskPaneVisible;
            config.Display.TaskPaneWidth   = _original.Display.TaskPaneWidth;
            config.Processing.MaxBodyLength       = (int)numMaxBody.Value;
            config.Processing.ConcurrentRequests     = _original.Processing.ConcurrentRequests;
            config.Processing.AutoAnalyzeNewMail     = chkAutoAnalyze.Checked;
            config.Processing.IncludeAttachmentNames = chkIncludeAttachments.Checked;

            var selectedStore = cmbStore.SelectedItem as StoreItem;
            config.Processing.TargetStoreId = selectedStore != null ? selectedStore.StoreId : "";

            // 발신자 규칙 수집 (R-01)
            config.Rules = new Models.RulesConfig();
            foreach (DataGridViewRow row in dgvRules.Rows)
            {
                if (row.IsNewRow) continue;
                string pattern = row.Cells["Pattern"].Value?.ToString() ?? "";
                if (string.IsNullOrWhiteSpace(pattern)) continue;

                bool enabled = row.Cells["Enabled"].Value is bool b && b;
                config.Rules.SenderRules.Add(new Models.SenderRule
                {
                    Enabled  = enabled,
                    Type     = row.Cells["Type"].Value?.ToString() ?? "email",
                    Pattern  = pattern.Trim(),
                    Priority = row.Cells["Priority"].Value?.ToString() ?? "normal",
                    Note     = row.Cells["Note"].Value?.ToString() ?? ""
                });
            }

            return config;
        }
    }
}
