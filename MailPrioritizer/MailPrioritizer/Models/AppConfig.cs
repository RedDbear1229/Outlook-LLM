using System.Collections.Generic;

namespace MailPrioritizer.Models
{
    public class AppConfig
    {
        public LlmConfig Llm { get; set; } = new LlmConfig();
        public ClassificationConfig Classification { get; set; } = new ClassificationConfig();
        public DisplayConfig Display { get; set; } = new DisplayConfig();
        public ProcessingConfig Processing { get; set; } = new ProcessingConfig();
        public RulesConfig Rules { get; set; } = new RulesConfig();
    }

    public class LlmConfig
    {
        public string Endpoint { get; set; } = "https://api.anthropic.com/v1";

        // 메모리 내에서는 항상 평문. 디스크 저장 시 ConfigManager가 DPAPI 암호화.
        public string ApiToken { get; set; } = "";

        public string ModelName { get; set; } = "claude-sonnet-4-20250514";
        public int MaxTokens { get; set; } = 1024;
        public int TimeoutSeconds { get; set; } = 30;

        // 빈 문자열이면 LlmService가 내장 기본 프롬프트를 사용한다.
        public string SystemPrompt { get; set; } = "";
    }

    public class ClassificationConfig
    {
        public bool AutoMoveToFolder { get; set; } = true;
        public string FolderPrefix { get; set; } = "우선순위";

        public Dictionary<string, string> FolderNames { get; set; } = new Dictionary<string, string>
        {
            { "urgent", "긴급" },
            { "high",   "높음" },
            { "normal", "보통" },
            { "low",    "낮음" }
        };

        public string GetFolderName(Priority priority)
        {
            string key = priority.ToString().ToLowerInvariant();
            string name;
            if (FolderNames.TryGetValue(key, out name))
                return name;
            return priority.ToKorean();
        }
    }

    public class DisplayConfig
    {
        public bool TagSubjectWithPriority { get; set; } = false;
        /// <summary>R-03: Outlook 재시작 시 Task Pane 표시 상태 복원.</summary>
        public bool TaskPaneVisible { get; set; } = true;
        /// <summary>R-03: Outlook 재시작 시 Task Pane 너비 복원.</summary>
        public int TaskPaneWidth { get; set; } = 320;
    }

    public class ProcessingConfig
    {
        public int MaxBodyLength { get; set; } = 4000;
        public int ConcurrentRequests { get; set; } = 3;

        /// <summary>
        /// 대상 저장소 ID. 빈 문자열이면 기본 저장소(Default Store) 사용.
        /// </summary>
        public string TargetStoreId { get; set; } = "";

        public bool AutoAnalyzeNewMail { get; set; } = false;
        /// <summary>R-04: 첨부파일명을 LLM 프롬프트에 포함할지 여부.</summary>
        public bool IncludeAttachmentNames { get; set; } = true;
    }

    // ──────────────────────────────────────────────────────────────
    // R-01: 발신자/도메인 기반 규칙 엔진
    // ──────────────────────────────────────────────────────────────

    /// <summary>발신자 기반 자동 분류 규칙 1건.</summary>
    public class SenderRule
    {
        /// <summary>"email" (정확 일치) 또는 "domain" (@이후 도메인 일치).</summary>
        public string Type { get; set; } = "email";
        /// <summary>매칭 패턴 (예: ceo@company.com 또는 newsletter.com).</summary>
        public string Pattern { get; set; } = "";
        /// <summary>적용할 우선순위 (urgent / high / normal / low).</summary>
        public string Priority { get; set; } = "normal";
        /// <summary>규칙 활성화 여부.</summary>
        public bool Enabled { get; set; } = true;
        /// <summary>관리용 메모 (LLM에 전달되지 않음).</summary>
        public string Note { get; set; } = "";
    }

    public class RulesConfig
    {
        public List<SenderRule> SenderRules { get; set; } = new List<SenderRule>();
    }
}
