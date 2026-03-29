using System.ComponentModel;

namespace MailPrioritizer.Models
{
    /// <summary>메일 우선순위 열거형. LLM이 반환하는 값과 1:1 매핑.</summary>
    public enum Priority
    {
        [Description("긴급")]
        Urgent = 1,

        [Description("높음")]
        High = 2,

        [Description("보통")]
        Normal = 3,

        [Description("낮음")]
        Low = 4
    }

    public static class PriorityExtensions
    {
        public static string ToKorean(this Priority priority)
        {
            switch (priority)
            {
                case Priority.Urgent: return "긴급";
                case Priority.High:   return "높음";
                case Priority.Normal: return "보통";
                case Priority.Low:    return "낮음";
                default:              return "보통";
            }
        }

        public static string ToEmoji(this Priority priority)
        {
            switch (priority)
            {
                case Priority.Urgent: return "🔴";
                case Priority.High:   return "🟠";
                case Priority.Normal: return "🟢";
                case Priority.Low:    return "⚪";
                default:              return "⚪";
            }
        }

        public static Priority FromString(string value)
        {
            if (value == null) return Priority.Normal;
            switch (value.ToLowerInvariant().Trim())
            {
                case "urgent": return Priority.Urgent;
                case "high":   return Priority.High;
                case "normal": return Priority.Normal;
                case "low":    return Priority.Low;
                default:       return Priority.Normal;
            }
        }
    }
}
