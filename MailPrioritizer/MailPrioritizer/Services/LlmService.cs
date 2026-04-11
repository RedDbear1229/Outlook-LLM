using System;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using MailPrioritizer.Models;
using MailPrioritizer.Utils;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace MailPrioritizer.Services
{
    /// <summary>
    /// LLM API 클라이언트.
    /// Endpoint URL에 "anthropic"이 포함되면 Claude Messages API,
    /// 그 외에는 OpenAI Chat Completions API 형식으로 호출한다.
    /// </summary>
    public class LlmService : IDisposable
    {
        private HttpClient _httpClient;
        private AppConfig _config;
        private SemaphoreSlim _throttle;
        private int _concurrentRequests;
        private bool _disposed;

        // R-13: Health check state
        public DateTime? LastSuccessAt { get; private set; }
        public DateTime? LastErrorAt   { get; private set; }
        public string   LastErrorMessage { get; private set; }

        internal static readonly string DefaultSystemPrompt =
            "당신은 이메일 분석 도우미입니다. 다음 이메일을 분석하여 반드시 아래 JSON 형식으로만 응답하세요.\n\n" +
            "**우선순위 기준:**\n" +
            "- urgent: 즉시 조치 필요 (마감 임박, 시스템 장애, 긴급 요청, 에스컬레이션)\n" +
            "- high: 당일 처리 필요 (중요 업무 요청, 의사결정 필요, 상위 보고)\n" +
            "- normal: 일반 업무 (정보 공유, 일반 질문, 진행 상황 공유)\n" +
            "- low: 참고용 (뉴스레터, 자동 알림, 광고, 공지사항)\n\n" +
            "**응답 형식 (JSON만 출력):**\n" +
            "{\"summary\":\"2-3문장 한국어 요약\",\"priority\":\"urgent|high|normal|low\",\"priority_reason\":\"판단 근거 1문장\"}";

        public LlmService(AppConfig config)
        {
            _config = config;
            _concurrentRequests = config.Processing.ConcurrentRequests;
            _throttle = new SemaphoreSlim(_concurrentRequests);
            _httpClient = CreateHttpClient(config);
        }

        public void ReloadConfig(AppConfig newConfig)
        {
            _config = newConfig;
            var oldClient = _httpClient;
            _httpClient = CreateHttpClient(newConfig);
            oldClient.Dispose();

            if (_concurrentRequests != newConfig.Processing.ConcurrentRequests)
            {
                _concurrentRequests = newConfig.Processing.ConcurrentRequests;
                var oldThrottle = _throttle;
                _throttle = new SemaphoreSlim(_concurrentRequests);
                oldThrottle.Dispose();
            }
        }

        private bool IsClaudeApi
        {
            get
            {
                return _config.Llm.Endpoint
                    .IndexOf("anthropic", StringComparison.OrdinalIgnoreCase) >= 0;
            }
        }

        private string SystemPrompt
        {
            get
            {
                return string.IsNullOrWhiteSpace(_config.Llm.SystemPrompt)
                    ? DefaultSystemPrompt
                    : _config.Llm.SystemPrompt;
            }
        }

        /// <summary>메일을 분석하여 요약과 우선순위를 반환한다.</summary>
        public async Task<MailAnalysis> AnalyzeMailAsync(
            string subject, string body, string sender,
            string attachments = "",
            CancellationToken cancellationToken = default(CancellationToken))
        {
            string truncSubject = subject != null && subject.Length > 50
                ? subject.Substring(0, 50) + "..." : subject ?? "";
            Logger.Info("AnalyzeMailAsync: start, subject=\"" + truncSubject + "\"");
            Logger.Debug("AnalyzeMailAsync: using " + (IsClaudeApi ? "Claude" : "OpenAI") + " API");

            await _throttle.WaitAsync(cancellationToken);
            try
            {
                string userContent = BuildUserMessage(subject, body, sender, attachments);
                string responseText = IsClaudeApi
                    ? await CallClaudeApiAsync(userContent, cancellationToken)
                    : await CallOpenAiApiAsync(userContent, cancellationToken);

                var result = ParseLlmResponse(responseText);
                result.ModelName = _config.Llm.ModelName;
                Logger.Info("AnalyzeMailAsync: complete, priority=" + result.Priority);
                LastSuccessAt = DateTime.Now;
                return result;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                Logger.Error("AnalyzeMailAsync: API call failed", ex);
                LastErrorAt = DateTime.Now;
                LastErrorMessage = ex.Message;
                return MailAnalysis.CreateFallback("API 호출 오류: " + ex.Message);
            }
            finally
            {
                _throttle.Release();
            }
        }

        /// <summary>연결 테스트. (bool success, string message) 반환.</summary>
        public async Task<Tuple<bool, string>> TestConnectionAsync(
            CancellationToken cancellationToken = default(CancellationToken))
        {
            Logger.Info("TestConnectionAsync: start");
            try
            {
                string testContent = "Hello. Please respond with: {\"summary\":\"test\",\"priority\":\"normal\",\"priority_reason\":\"test\"}";
                string responseText = IsClaudeApi
                    ? await CallClaudeApiAsync(testContent, cancellationToken)
                    : await CallOpenAiApiAsync(testContent, cancellationToken);

                if (!string.IsNullOrEmpty(responseText))
                {
                    Logger.Info("TestConnectionAsync: success");
                    return Tuple.Create(true, "연결 성공! 모델: " + _config.Llm.ModelName);
                }
                else
                {
                    Logger.Warn("TestConnectionAsync: empty response");
                    return Tuple.Create(false, "응답이 비어 있습니다.");
                }
            }
            catch (HttpRequestException ex)
            {
                Logger.Error("TestConnectionAsync: HTTP error", ex);
                return Tuple.Create(false, "HTTP 오류: " + ex.Message);
            }
            catch (TaskCanceledException)
            {
                Logger.Warn("TestConnectionAsync: timeout");
                return Tuple.Create(false, "연결 시간 초과 (" + _config.Llm.TimeoutSeconds + "초)");
            }
            catch (Exception ex)
            {
                Logger.Error("TestConnectionAsync: unexpected error", ex);
                return Tuple.Create(false, "오류: " + ex.Message);
            }
        }

        // ──────────────────────────────────────────────────────────────
        // Claude Messages API
        // POST {endpoint}/messages
        // ──────────────────────────────────────────────────────────────
        private async Task<string> CallClaudeApiAsync(
            string userContent, CancellationToken ct)
        {
            string url = _config.Llm.Endpoint.TrimEnd('/') + "/messages";

            var requestBody = new
            {
                model = _config.Llm.ModelName,
                max_tokens = _config.Llm.MaxTokens,
                system = SystemPrompt,
                messages = new[]
                {
                    new { role = "user", content = userContent }
                }
            };

            string json = JsonConvert.SerializeObject(requestBody);
            string token = _config.Llm.ApiToken;
            Func<HttpRequestMessage> factory = () =>
            {
                var req = new HttpRequestMessage(HttpMethod.Post, url);
                req.Headers.Add("x-api-key", token);
                req.Headers.Add("anthropic-version", "2023-06-01");
                req.Content = new StringContent(json, Encoding.UTF8, "application/json");
                return req;
            };

            // 429 Rate Limit 시 1회 재시도
            using (var response = await SendWithRetryAsync(factory, ct))
            {
                response.EnsureSuccessStatusCode();

                string responseJson = await response.Content.ReadAsStringAsync();
                var jObj = JObject.Parse(responseJson);

                // 응답: { "content": [{ "type": "text", "text": "..." }] }
                return jObj["content"]?[0]?["text"]?.ToString() ?? "";
            }
        }

        // ──────────────────────────────────────────────────────────────
        // OpenAI Chat Completions API
        // POST {endpoint}/chat/completions
        // ──────────────────────────────────────────────────────────────
        private async Task<string> CallOpenAiApiAsync(
            string userContent, CancellationToken ct)
        {
            string url = _config.Llm.Endpoint.TrimEnd('/') + "/chat/completions";

            var requestBody = new
            {
                model = _config.Llm.ModelName,
                max_tokens = _config.Llm.MaxTokens,
                messages = new[]
                {
                    new { role = "system", content = SystemPrompt },
                    new { role = "user",   content = userContent }
                }
            };

            string json = JsonConvert.SerializeObject(requestBody);
            string token = _config.Llm.ApiToken;
            Func<HttpRequestMessage> factory = () =>
            {
                var req = new HttpRequestMessage(HttpMethod.Post, url);
                req.Headers.Add("Authorization", "Bearer " + token);
                req.Content = new StringContent(json, Encoding.UTF8, "application/json");
                return req;
            };

            using (var response = await SendWithRetryAsync(factory, ct))
            {
                response.EnsureSuccessStatusCode();

                string responseJson = await response.Content.ReadAsStringAsync();
                var jObj = JObject.Parse(responseJson);

                // 응답: { "choices": [{ "message": { "content": "..." } }] }
                return jObj["choices"]?[0]?["message"]?["content"]?.ToString() ?? "";
            }
        }

        /// <summary>429/5xx 시 지수 백오프로 최대 3회 재시도. factory로 매번 새 HttpRequestMessage 생성.</summary>
        private async Task<HttpResponseMessage> SendWithRetryAsync(
            Func<HttpRequestMessage> factory, CancellationToken ct)
        {
            int maxRetries = 3;
            int delayMs = 2000; // 2초, 4초, 8초

            HttpResponseMessage response = null;
            for (int attempt = 0; attempt <= maxRetries; attempt++)
            {
                response?.Dispose();
                response = null;

                response = await _httpClient.SendAsync(factory(), ct);
                int statusCode = (int)response.StatusCode;

                bool isRetryable = statusCode == 429 || statusCode == 500
                                || statusCode == 502 || statusCode == 503;

                if (!isRetryable || attempt == maxRetries)
                    break;

                Logger.Warn(string.Format(
                    "SendWithRetryAsync: HTTP {0}, retry {1}/{2} after {3}ms",
                    statusCode, attempt + 1, maxRetries, delayMs));

                await Task.Delay(delayMs, ct);
                delayMs *= 2;
            }

            return response;
        }

        // ──────────────────────────────────────────────────────────────
        // LLM 응답 JSON 파싱 (공통)
        // ──────────────────────────────────────────────────────────────
        private static MailAnalysis ParseLlmResponse(string responseText)
        {
            if (string.IsNullOrWhiteSpace(responseText))
                return MailAnalysis.CreateFallback("LLM 응답이 비어 있습니다.");

            // JSON 블록 추출 시도 (```json ... ``` 감싸여 있을 수 있음)
            string jsonText = ExtractJson(responseText);

            try
            {
                var jObj = JObject.Parse(jsonText);
                string summary = jObj["summary"]?.ToString() ?? "(요약 없음)";
                string priorityStr = jObj["priority"]?.ToString() ?? "normal";
                string reason = jObj["priority_reason"]?.ToString() ?? "";

                return new MailAnalysis
                {
                    Summary = summary,
                    Priority = PriorityExtensions.FromString(priorityStr),
                    PriorityReason = reason,
                    AnalyzedAt = DateTime.Now,
                    IsFallback = false
                };
            }
            catch (Exception ex)
            {
                string snippet = responseText.Substring(0, Math.Min(100, responseText.Length));
                Logger.Error("ParseLlmResponse: JSON parse failed, response snippet: " + snippet, ex);
                return MailAnalysis.CreateFallback("응답 파싱 실패: " + snippet);
            }
        }

        private static string ExtractJson(string text)
        {
            if (string.IsNullOrEmpty(text)) return text ?? "";

            // ```json ... ``` 또는 ``` ... ``` 제거
            var fenceMatch = Regex.Match(text, @"```(?:json)?\s*([\s\S]*?)```");
            if (fenceMatch.Success)
                return fenceMatch.Groups[1].Value.Trim();

            // { ... } 블록 추출
            int start = text.IndexOf('{');
            int end = text.LastIndexOf('}');
            if (start >= 0 && end > start)
                return text.Substring(start, end - start + 1);

            return text.Trim();
        }

        private string BuildUserMessage(string subject, string body, string sender, string attachments = "")
        {
            int maxLen = _config.Processing.MaxBodyLength;
            string truncatedBody = body != null && body.Length > maxLen
                ? body.Substring(0, maxLen) + "\n...(이하 생략)"
                : (body ?? "");

            if (!string.IsNullOrEmpty(attachments))
                return string.Format(
                    "발신자: {0}\n제목: {1}\n첨부파일: {2}\n\n본문:\n{3}",
                    sender ?? "", subject ?? "", attachments, truncatedBody);

            return string.Format(
                "발신자: {0}\n제목: {1}\n\n본문:\n{2}",
                sender ?? "", subject ?? "", truncatedBody);
        }

        private static HttpClient CreateHttpClient(AppConfig config)
        {
            return new HttpClient
            {
                Timeout = TimeSpan.FromSeconds(config.Llm.TimeoutSeconds)
            };
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            _httpClient?.Dispose();
            _throttle?.Dispose();
        }
    }
}
