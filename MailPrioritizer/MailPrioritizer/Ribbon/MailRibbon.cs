using System;
using System.Windows.Forms;
using MailPrioritizer.Utils;
using Office = Microsoft.Office.Core;

namespace MailPrioritizer.Ribbon
{
    public partial class MailRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        public MailRibbon() { }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        /// <summary>API 미설정 시 안내 메시지 표시 후 설정 창 열기. 설정 완료 여부 반환.</summary>
        private bool EnsureConfigured(Office.IRibbonControl control)
        {
            if (Globals.ThisAddIn.IsConfigured()) return true;
            MessageBox.Show(
                "API 설정이 필요합니다.\n환경설정에서 Endpoint URL, API Token, 모델 이름을 입력해 주세요.",
                "MailPrioritizer",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            OnSettingsClick(control);
            return false;
        }

        // ──────────────────────────────────────────────────────────────
        // 환경설정
        // ──────────────────────────────────────────────────────────────
        public void OnSettingsClick(Office.IRibbonControl control)
        {
            using (var form = new Forms.SettingsForm(Globals.ThisAddIn.Config))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    Globals.ThisAddIn.ApplyConfigChange(form.ResultConfig, form.PlainToken);
                }
            }
        }

        // ──────────────────────────────────────────────────────────────
        // 선택 메일 요약
        // ──────────────────────────────────────────────────────────────
        public async void OnAnalyzeSelectedClick(Office.IRibbonControl control)
        {
            if (!EnsureConfigured(control)) return;

            Microsoft.Office.Interop.Outlook.Explorer explorer = null;
            Microsoft.Office.Interop.Outlook.Selection selection = null;
            Microsoft.Office.Interop.Outlook.MailItem mail = null;
            try
            {
                explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                if (explorer == null) return;
                selection = explorer.Selection;
                if (selection.Count == 0)
                {
                    MessageBox.Show("분석할 메일을 선택해 주세요.", "MailPrioritizer",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                mail = selection[1] as Microsoft.Office.Interop.Outlook.MailItem;
                if (mail == null)
                {
                    MessageBox.Show("메일 항목을 선택해 주세요.", "MailPrioritizer",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                Globals.ThisAddIn.SummaryControl.ShowAnalyzing();

                var analysis = await Globals.ThisAddIn.MailProcessor.AnalyzeSingleAsync(mail);
                Globals.ThisAddIn.SummaryControl.DisplayAnalysis(mail, analysis);
            }
            catch (Exception ex)
            {
                Logger.Error("OnAnalyzeSelectedClick: analysis failed", ex);
                MessageBox.Show("분석 오류: " + ex.Message, "MailPrioritizer",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                Globals.ThisAddIn.SummaryControl.ShowError(ex.Message);
            }
            finally
            {
                ComHelper.ReleaseAll(mail, selection, explorer);
            }
        }

        // ──────────────────────────────────────────────────────────────
        // 받은편지함 전체 분류
        // ──────────────────────────────────────────────────────────────
        public async void OnAnalyzeAllClick(Office.IRibbonControl control)
        {
            if (!EnsureConfigured(control)) return;

            using (var progressForm = new Forms.ProgressForm())
            {
                progressForm.Show();

                var cts = progressForm.CancellationTokenSource;
                var progress = new Progress<Services.MailProcessor.BatchProgress>(p =>
                {
                    if (!progressForm.IsDisposed)
                        progressForm.UpdateProgress(p);
                });

                try
                {
                    var result = await Globals.ThisAddIn.MailProcessor.AnalyzeInboxAsync(
                        Globals.ThisAddIn.Application, progress, cts.Token);

                    string summary = string.Format(
                        "분석 완료!\n\n긴급: {0}건\n높음: {1}건\n보통: {2}건\n낮음: {3}건\n실패: {4}건",
                        result.Urgent, result.High, result.Normal, result.Low, result.Failed);

                    if (result.FailedSubjects.Count > 0)
                    {
                        summary += "\n\n── 실패 메일 ──";
                        int showCount = Math.Min(result.FailedSubjects.Count, 10);
                        for (int i = 0; i < showCount; i++)
                        {
                            string subj = result.FailedSubjects[i];
                            if (subj.Length > 40) subj = subj.Substring(0, 40) + "...";
                            summary += "\n• " + subj;
                        }
                        if (result.FailedSubjects.Count > 10)
                            summary += string.Format("\n  ...외 {0}건", result.FailedSubjects.Count - 10);
                    }

                    progressForm.Close();
                    MessageBox.Show(summary,
                        "MailPrioritizer — 완료",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                catch (OperationCanceledException)
                {
                    progressForm.Close();
                    MessageBox.Show("분석이 취소되었습니다.\n처리된 메일은 유지됩니다.",
                        "MailPrioritizer", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    Logger.Error("OnAnalyzeAllClick: batch analysis failed", ex);
                    progressForm.Close();
                    MessageBox.Show("오류 발생: " + ex.Message, "MailPrioritizer",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        // ──────────────────────────────────────────────────────────────
        // 전체 재분석
        // ──────────────────────────────────────────────────────────────
        public void OnReanalyzeAllClick(Office.IRibbonControl control)
        {
            if (!EnsureConfigured(control)) return;

            var answer = MessageBox.Show(
                "받은편지함의 모든 메일 분석 플래그를 초기화하고 전체 재분석을 실행합니다.\n계속하시겠습니까?",
                "MailPrioritizer — 전체 재분석",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);
            if (answer != DialogResult.Yes) return;

            int resetCount = Globals.ThisAddIn.MailProcessor.ResetAllAnalysisFlags(
                Globals.ThisAddIn.Application);
            Logger.Info("OnReanalyzeAllClick: reset " + resetCount + " mails, starting batch analysis");

            // 리셋 후 일괄 분석 실행 (OnAnalyzeAllClick과 동일 흐름)
            OnAnalyzeAllClick(control);
        }

        // ──────────────────────────────────────────────────────────────
        // 결과 내보내기
        // ──────────────────────────────────────────────────────────────
        public void OnExportResultsClick(Office.IRibbonControl control)
        {
            using (var dialog = new SaveFileDialog())
            {
                dialog.Title = "분석 결과 내보내기";
                dialog.Filter = "CSV 파일 (*.csv)|*.csv";
                dialog.FileName = "MailPrioritizer_Export_" + DateTime.Now.ToString("yyyyMMdd") + ".csv";

                if (dialog.ShowDialog() != DialogResult.OK) return;

                try
                {
                    int count = Globals.ThisAddIn.MailProcessor.ExportAnalyzedMails(
                        Globals.ThisAddIn.Application, dialog.FileName);
                    MessageBox.Show(
                        string.Format("{0}건의 분석 결과를 내보냈습니다.\n{1}", count, dialog.FileName),
                        "MailPrioritizer",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    Logger.Error("OnExportResultsClick: export failed", ex);
                    MessageBox.Show("내보내기 오류: " + ex.Message, "MailPrioritizer",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}
