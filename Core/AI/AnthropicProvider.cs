using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace SpreadsheetApp.Core.AI
{
    public sealed class AnthropicProvider : IInferenceProvider
    {
        private readonly HttpClient _http;
        private readonly string _model;

        public AnthropicProvider(string apiKey, string? model = null, HttpClient? http = null)
        {
            if (string.IsNullOrWhiteSpace(apiKey)) throw new ArgumentNullException(nameof(apiKey));
            _http = http ?? new HttpClient { Timeout = TimeSpan.FromSeconds(20) };
            _http.DefaultRequestHeaders.Add("x-api-key", apiKey);
            _http.DefaultRequestHeaders.Add("anthropic-version", "2023-06-01");
            _model = string.IsNullOrWhiteSpace(model) ? Environment.GetEnvironmentVariable("ANTHROPIC_MODEL") ?? "claude-3-haiku-20240307" : model;
        }

        public async Task<AIResult> GenerateFillAsync(AIContext context, CancellationToken cancellationToken)
        {
            string endpoint = Environment.GetEnvironmentVariable("ANTHROPIC_BASE_URL") ?? "https://api.anthropic.com/v1/messages";
            var prompt = "You are a spreadsheet fill assistant. Respond ONLY with strict JSON: {\"cells\": [[...]]} sized exactly to rows x cols; plain text strings; no formatting, no code blocks, no extra keys.\n\n" +
                         $"Sheet={context.SheetName}; Title={(context.Title ?? "").Trim()}; Start=({context.StartRow+1},{context.StartCol+1}); Rows={context.Rows}; Cols={context.Cols}; Prompt={(context.Prompt ?? "").Trim()}";
            int maxTokens = 2048;
            try { var s = Environment.GetEnvironmentVariable("ANTHROPIC_MAX_TOKENS"); if (!string.IsNullOrWhiteSpace(s)) maxTokens = int.Parse(s); } catch { }
            var body = new
            {
                model = _model,
                max_tokens = maxTokens,
                temperature = 0.2,
                messages = new object[]
                {
                    new { role = "user", content = new object[] { new { type = "text", text = prompt } } }
                }
            };
            using var req = new HttpRequestMessage(HttpMethod.Post, endpoint)
            {
                Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json")
            };
            var t0 = DateTime.UtcNow;
            using var resp = await _http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false);
            resp.EnsureSuccessStatusCode();
            await using var stream = await resp.Content.ReadAsStreamAsync(cancellationToken).ConfigureAwait(false);
            using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken).ConfigureAwait(false);
            var result = new AIResult { Provider = "Anthropic", Model = _model, LatencyMs = (int)(DateTime.UtcNow - t0).TotalMilliseconds };
            try
            {
                if (doc.RootElement.TryGetProperty("model", out var mdl) && mdl.ValueKind == JsonValueKind.String) result.Model = mdl.GetString();
                if (doc.RootElement.TryGetProperty("usage", out var usage) && usage.ValueKind == JsonValueKind.Object)
                {
                    var u = new AIUsage();
                    if (usage.TryGetProperty("input_tokens", out var it) && it.ValueKind == JsonValueKind.Number) u.InputTokens = it.GetInt32();
                    if (usage.TryGetProperty("output_tokens", out var ot) && ot.ValueKind == JsonValueKind.Number) u.OutputTokens = ot.GetInt32();
                    if (u.InputTokens.HasValue && u.OutputTokens.HasValue) u.TotalTokens = u.InputTokens + u.OutputTokens;
                    try { var lim = Environment.GetEnvironmentVariable("ANTHROPIC_CONTEXT_TOKENS"); if (!string.IsNullOrWhiteSpace(lim)) u.ContextLimit = int.Parse(lim); } catch { }
                    result.Usage = u;
                }
            }
            catch { }
            if (doc.RootElement.TryGetProperty("content", out var contentEl) && contentEl.ValueKind == JsonValueKind.Array && contentEl.GetArrayLength() > 0)
            {
                string? text = null;
                foreach (var part in contentEl.EnumerateArray())
                {
                    if (part.TryGetProperty("text", out var t) && t.ValueKind == JsonValueKind.String) { text = t.GetString(); break; }
                }
                if (!string.IsNullOrEmpty(text))
                {
                    try
                    {
                        var cellsDoc = JsonDocument.Parse(text);
                        if (cellsDoc.RootElement.ValueKind == JsonValueKind.Object && cellsDoc.RootElement.TryGetProperty("cells", out var cellsEl))
                        {
                            var rows = ExternalApiProvider.ParseCellsArray(cellsEl);
                            result.Cells = ExternalApiProvider.FitToShape(rows, context.Rows, context.Cols);
                            return result;
                        }
                    }
                    catch { }
                    var lines = text.Replace("\r", string.Empty).Split('\n');
                    result.Cells = ExternalApiProvider.FitToShape(lines, context.Rows, context.Cols);
                    return result;
                }
            }
            result.Cells = ExternalApiProvider.FitToShape(Array.Empty<string>(), context.Rows, context.Cols);
            return result;
        }
    }
}
