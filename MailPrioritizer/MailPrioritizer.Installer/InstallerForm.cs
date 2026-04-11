using System;
using System.Drawing;
using System.Windows.Forms;

namespace MailPrioritizer.Installer
{
    internal class InstallerForm : Form
    {
        // ── 컨트롤 ──────────────────────────────────────────────────────────
        private Panel   _panelStatus;
        private Label   _lblStatusIcon;
        private Label   _lblStatusText;
        private Label   _lblDllPath;
        private ListBox _lstLog;
        private Button  _btnInstall;
        private Button  _btnRepair;
        private Button  _btnUninstall;
        private Button  _btnClose;

        private string _dllDirectory;

        // ── 생성자 ──────────────────────────────────────────────────────────
        public InstallerForm()
        {
            BuildUi();
            RefreshStatus();
        }

        // ── UI 구성 (Designer 없이 코드로 생성) ─────────────────────────────
        private void BuildUi()
        {
            SuspendLayout();

            Text            = "MailPrioritizer 설치 관리자";
            Size            = new Size(540, 420);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox     = false;
            MinimizeBox     = false;
            StartPosition   = FormStartPosition.CenterScreen;
            Font            = new Font("맑은 고딕", 9f);

            // ── 상태 패널 ──────────────────────────────────────────────────
            _panelStatus = new Panel
            {
                Location    = new Point(12, 12),
                Size        = new Size(500, 72),
                BorderStyle = BorderStyle.FixedSingle,
                BackColor   = Color.FromArgb(248, 248, 248)
            };

            _lblStatusIcon = new Label
            {
                Location  = new Point(10, 10),
                Size      = new Size(28, 28),
                Font      = new Font("맑은 고딕", 14f, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter
            };

            _lblStatusText = new Label
            {
                Location  = new Point(44, 10),
                Size      = new Size(448, 22),
                Font      = new Font("맑은 고딕", 10f, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft
            };

            _lblDllPath = new Label
            {
                Location  = new Point(44, 36),
                Size      = new Size(448, 18),
                ForeColor = Color.DimGray,
                TextAlign = ContentAlignment.MiddleLeft
            };

            _panelStatus.Controls.Add(_lblStatusIcon);
            _panelStatus.Controls.Add(_lblStatusText);
            _panelStatus.Controls.Add(_lblDllPath);

            // ── 로그 영역 ──────────────────────────────────────────────────
            var lblLog = new Label
            {
                Text      = "작업 로그",
                Location  = new Point(12, 96),
                AutoSize  = true,
                ForeColor = Color.DimGray
            };

            _lstLog = new ListBox
            {
                Location            = new Point(12, 114),
                Size                = new Size(500, 188),
                Font                = new Font("Consolas", 8.5f),
                HorizontalScrollbar = true,
                SelectionMode       = SelectionMode.None
            };

            // ── 버튼 ───────────────────────────────────────────────────────
            _btnInstall = new Button
            {
                Text     = "설치",
                Location = new Point(12, 316),
                Size     = new Size(108, 36),
                UseVisualStyleBackColor = true
            };
            _btnInstall.Click += OnInstallClick;

            _btnRepair = new Button
            {
                Text     = "복구",
                Location = new Point(128, 316),
                Size     = new Size(108, 36),
                UseVisualStyleBackColor = true
            };
            _btnRepair.Click += OnRepairClick;

            _btnUninstall = new Button
            {
                Text      = "제거",
                Location  = new Point(244, 316),
                Size      = new Size(108, 36),
                ForeColor = Color.Firebrick,
                UseVisualStyleBackColor = true
            };
            _btnUninstall.Click += OnUninstallClick;

            _btnClose = new Button
            {
                Text     = "닫기",
                Location = new Point(404, 316),
                Size     = new Size(108, 36),
                UseVisualStyleBackColor = true
            };
            _btnClose.Click += (s, e) => Close();

            Controls.AddRange(new Control[]
            {
                _panelStatus, lblLog, _lstLog,
                _btnInstall, _btnRepair, _btnUninstall, _btnClose
            });

            ResumeLayout(false);
        }

        // ── 상태 갱신 ────────────────────────────────────────────────────────
        private void RefreshStatus()
        {
            _dllDirectory = RegistrationHelper.FindDllDirectory();

            string details;
            RegistrationHelper.Status status = RegistrationHelper.GetStatus(out details);

            switch (status)
            {
                case RegistrationHelper.Status.NotRegistered:
                    SetStatus("●", "미등록", Color.Firebrick);
                    EnableButtons(install: _dllDirectory != null, repair: false, uninstall: false);
                    break;

                case RegistrationHelper.Status.RegisteredOk:
                    SetStatus("●", "정상 등록됨", Color.DarkGreen);
                    EnableButtons(install: false, repair: true, uninstall: true);
                    break;

                case RegistrationHelper.Status.LoadBehaviorDisabled:
                    SetStatus("●", "비활성됨 — Outlook이 Add-in을 중지했습니다", Color.DarkOrange);
                    EnableButtons(install: false, repair: _dllDirectory != null, uninstall: true);
                    break;

                case RegistrationHelper.Status.ManifestMissing:
                    SetStatus("●", "Manifest 파일 없음 — 복구가 필요합니다", Color.DarkOrange);
                    EnableButtons(install: false, repair: _dllDirectory != null, uninstall: true);
                    break;
            }

            if (_dllDirectory != null)
                _lblDllPath.Text = "DLL 경로: " + _dllDirectory;
            else
                _lblDllPath.Text = "MailPrioritizer.dll을 찾을 수 없습니다. 먼저 빌드하세요.";
        }

        private void SetStatus(string icon, string text, Color color)
        {
            _lblStatusIcon.Text      = icon;
            _lblStatusIcon.ForeColor = color;
            _lblStatusText.Text      = "상태: " + text;
            _lblStatusText.ForeColor = color;
            _panelStatus.BackColor   = Color.FromArgb(
                Math.Min(255, color.R + 200),
                Math.Min(255, color.G + 200),
                Math.Min(255, color.B + 200));
        }

        private void EnableButtons(bool install, bool repair, bool uninstall)
        {
            _btnInstall.Enabled   = install;
            _btnRepair.Enabled    = repair;
            _btnUninstall.Enabled = uninstall;
        }

        // ── 로그 출력 ────────────────────────────────────────────────────────
        private void Log(string message)
        {
            _lstLog.Items.Add("  " + message);
            _lstLog.SelectedIndex = _lstLog.Items.Count - 1;
        }

        // ── Outlook 실행 중 경고 ─────────────────────────────────────────────
        private bool ConfirmIfOutlookRunning()
        {
            if (!RegistrationHelper.IsOutlookRunning()) return true;
            DialogResult ans = MessageBox.Show(
                "Outlook이 현재 실행 중입니다.\n\n변경 사항은 Outlook 재시작 후에 반영됩니다.\n계속하시겠습니까?",
                "MailPrioritizer 설치 관리자",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);
            return ans == DialogResult.Yes;
        }

        // ── 이벤트 핸들러 ────────────────────────────────────────────────────
        private void OnInstallClick(object sender, EventArgs e)
        {
            if (!ConfirmIfOutlookRunning()) return;
            _lstLog.Items.Clear();
            try
            {
                RegistrationHelper.Install(_dllDirectory, Log);
                RefreshStatus();
                Log(string.Empty);
                Log("Outlook을 재시작하면 '메일 분석' 리본 탭이 나타납니다.");
                MessageBox.Show(
                    "설치가 완료되었습니다.\nOutlook을 재시작하세요.",
                    "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Log("[오류] " + ex.Message);
                MessageBox.Show("설치 중 오류:\n" + ex.Message, "오류",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OnRepairClick(object sender, EventArgs e)
        {
            if (!ConfirmIfOutlookRunning()) return;
            _lstLog.Items.Clear();
            try
            {
                RegistrationHelper.Repair(_dllDirectory, Log);
                RefreshStatus();
                Log(string.Empty);
                Log("Outlook을 재시작하면 변경 사항이 반영됩니다.");
                MessageBox.Show(
                    "복구가 완료되었습니다.\nOutlook을 재시작하세요.",
                    "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Log("[오류] " + ex.Message);
                MessageBox.Show("복구 중 오류:\n" + ex.Message, "오류",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OnUninstallClick(object sender, EventArgs e)
        {
            DialogResult confirm = MessageBox.Show(
                "MailPrioritizer를 Outlook에서 제거하시겠습니까?",
                "제거 확인", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (confirm != DialogResult.Yes) return;

            if (!ConfirmIfOutlookRunning()) return;
            _lstLog.Items.Clear();
            try
            {
                RegistrationHelper.Uninstall(Log);
                RefreshStatus();
                Log(string.Empty);
                Log("Outlook을 재시작하면 리본 탭이 사라집니다.");
                MessageBox.Show(
                    "제거가 완료되었습니다.\nOutlook을 재시작하세요.",
                    "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Log("[오류] " + ex.Message);
                MessageBox.Show("제거 중 오류:\n" + ex.Message, "오류",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
