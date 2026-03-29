using System;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using MailPrioritizer.Services;

namespace MailPrioritizer.Forms
{
    /// <summary>일괄 분석 진행률 표시 폼.</summary>
    public class ProgressForm : Form
    {
        public CancellationTokenSource CancellationTokenSource { get; } = new CancellationTokenSource();

        private ProgressBar _progressBar;
        private Label _lblStatus;
        private Label _lblCurrent;
        private Button _btnCancel;

        public ProgressForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "메일 분석 중...";
            this.Size = new Size(460, 180);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Font = new Font("맑은 고딕", 9f);

            // 닫기 버튼으로 취소
            this.FormClosing += (s, e) =>
            {
                if (!CancellationTokenSource.IsCancellationRequested)
                    CancellationTokenSource.Cancel();
            };

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(16)
            };

            _lblStatus = new Label
            {
                Text = "분석 준비 중...",
                Dock = DockStyle.Fill,
                AutoSize = false,
                Height = 24
            };

            _progressBar = new ProgressBar
            {
                Dock = DockStyle.Fill,
                Height = 20,
                Style = ProgressBarStyle.Blocks
            };

            _lblCurrent = new Label
            {
                Text = "",
                Dock = DockStyle.Fill,
                ForeColor = Color.Gray,
                AutoSize = false,
                Height = 22
            };

            _btnCancel = new Button
            {
                Text = "취소",
                Width = 80,
                Height = 28,
                Anchor = AnchorStyles.Right
            };
            _btnCancel.Click += (s, e) =>
            {
                CancellationTokenSource.Cancel();
                _btnCancel.Enabled = false;
                _btnCancel.Text = "취소 중...";
            };

            layout.Controls.Add(_lblStatus);
            layout.Controls.Add(_progressBar);
            layout.Controls.Add(_lblCurrent);
            layout.Controls.Add(_btnCancel);
            this.Controls.Add(layout);
        }

        public void UpdateProgress(MailProcessor.BatchProgress progress)
        {
            if (IsDisposed || !IsHandleCreated) return;

            if (InvokeRequired)
            {
                Invoke(new Action<MailProcessor.BatchProgress>(UpdateProgress), progress);
                return;
            }

            if (progress.Total > 0)
            {
                _progressBar.Maximum = progress.Total;
                _progressBar.Value   = Math.Min(progress.Processed, progress.Total);
                _lblStatus.Text = string.Format(
                    "{0} / {1} 메일 분석 완료  (긴급:{2} 높음:{3} 보통:{4} 낮음:{5} 실패:{6})",
                    progress.Processed, progress.Total,
                    progress.Urgent, progress.High, progress.Normal, progress.Low, progress.Failed);
            }
            else
            {
                _progressBar.Style = ProgressBarStyle.Marquee;
                _lblStatus.Text = "메일 목록 수집 중...";
            }

            if (!string.IsNullOrEmpty(progress.CurrentSubject))
                _lblCurrent.Text = "현재: " + Truncate(progress.CurrentSubject, 60);
        }

        private static string Truncate(string s, int max)
        {
            return s.Length <= max ? s : s.Substring(0, max) + "...";
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                CancellationTokenSource?.Dispose();
            base.Dispose(disposing);
        }
    }
}
