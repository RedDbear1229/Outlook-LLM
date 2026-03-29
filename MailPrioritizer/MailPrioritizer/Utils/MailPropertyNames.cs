namespace MailPrioritizer.Utils
{
    /// <summary>Outlook UserProperty 이름 상수. "LLM_*" 매직 스트링을 한 곳에서 관리.</summary>
    internal static class MailPropertyNames
    {
        public const string Priority       = "LLM_Priority";
        public const string Summary        = "LLM_Summary";
        public const string PriorityReason = "LLM_PriorityReason";
        public const string Analyzed       = "LLM_Analyzed";
    }
}
