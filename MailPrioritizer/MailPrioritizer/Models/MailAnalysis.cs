using System;

namespace MailPrioritizer.Models
{
    /// <summary>LLM 분석 결과 모델.</summary>
    public class MailAnalysis
    {
        /// <summary>2-3문장 한국어 요약.</summary>
        public string Summary { get; set; }

        /// <summary>우선순위.</summary>
        public Priority Priority { get; set; }

        /// <summary>우선순위 판단 근거 1문장.</summary>
        public string PriorityReason { get; set; }

        /// <summary>분석 시각.</summary>
        public DateTime AnalyzedAt { get; set; }

        /// <summary>분석 실패 여부 (API 오류 등).</summary>
        public bool IsFallback { get; set; }

        /// <summary>분석 실패 시 오류 메시지.</summary>
        public string ErrorMessage { get; set; }

        public static MailAnalysis CreateFallback(string errorMessage)
        {
            return new MailAnalysis
            {
                Summary = "(분석 실패)",
                Priority = Priority.Normal,
                PriorityReason = errorMessage,
                AnalyzedAt = DateTime.Now,
                IsFallback = true,
                ErrorMessage = errorMessage
            };
        }
    }
}
