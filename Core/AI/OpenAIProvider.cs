using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace SpreadsheetApp.Core.AI
{
    public sealed class OpenAIProvider : IInferenceProvider
    {
        private readonly HttpClient _http;
        private readonly string _model;

        public OpenAIProvider(string apiKey, string? model = null, HttpClient? http = null)
        {
            if (string.IsNullOrWhiteSpace(apiKey)) throw new ArgumentNullException(nameof(apiKey));
            _http = http ?? new HttpClient { Timeout = TimeSpan.FromSeconds(20) };
            _http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
            _http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            _model = string.IsNullOrWhiteSpace(model) ? Environment.GetEnvironmentVariable("OPENAI_MODEL") ?? "gpt-4o-mini" : model;
        }

        public async Task<AIResult> GenerateFillAsync(AIContext context, CancellationToken cancellationToken)
        {
            string endpoint = Environment.GetEnvironmentVariable("OPENAI_BASE_URL") ?? "https://api.openai.com/v1/chat/completions";
            var sys = "You are a spreadsheet fill assistant. Respond ONLY with strict JSON: {\"cells\": [[...]]} sized exactly to rows x cols; plain text strings; no formatting, no code blocks, no extra keys.";
            var user = $"Sheet={context.SheetName}; Title={(context.Title ?? "").Trim()}; Start=({context.StartRow+1},{context.StartCol+1}); Rows={context.Rows}; Cols={context.Cols}; Prompt={(context.Prompt ?? "").Trim()}";
            var body = new
            {
                model = _model,
                temperature = 0.2,
                messages = new object[]
                {
                    new { role = "system", content = sys },
                    new { role = "user", content = user }
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
            var result = new AIResult { Provider = "OpenAI", Model = _model, LatencyMs = (int)(DateTime.UtcNow - t0).TotalMilliseconds };
            try
            {
                if (doc.RootElement.TryGetProperty("model", out var mdl) && mdl.ValueKind == JsonValueKind.String) result.Model = mdl.GetString();
                if (doc.RootElement.TryGetProperty("usage", out var usage) && usage.ValueKind == JsonValueKind.Object)
                {
                    var u = new AIUsage();
                    if (usage.TryGetProperty("prompt_tokens", out var pt) && pt.ValueKind == JsonValueKind.Number) u.InputTokens = pt.GetInt32();
                    if (usage.TryGetProperty("completion_tokens", out var ctok) && ctok.ValueKind == JsonValueKind.Number) u.OutputTokens = ctok.GetInt32();
                    if (usage.TryGetProperty("total_tokens", out var tt) && tt.ValueKind == JsonValueKind.Number) u.TotalTokens = tt.GetInt32();
                    try { var lim = Environment.GetEnvironmentVariable("OPENAI_CONTEXT_TOKENS"); if (!string.IsNullOrWhiteSpace(lim)) u.ContextLimit = int.Parse(lim); } catch { }
                    result.Usage = u;
                }
            }
            catch { }
            // Try choices[0].message.content (string or array)
            if (doc.RootElement.TryGetProperty("choices", out var choices) && choices.ValueKind == JsonValueKind.Array && choices.GetArrayLength() > 0)
            {
                var msg = choices[0].GetProperty("message");
                string? content = null;
                if (msg.TryGetProperty("content", out var cnt))
                {
                    if (cnt.ValueKind == JsonValueKind.String) content = cnt.GetString();
                    else if (cnt.ValueKind == JsonValueKind.Array)
                    {
                        var sb = new StringBuilder();
                        foreach (var part in cnt.EnumerateArray())
                        {
                            if (part.TryGetProperty("text", out var t) && t.ValueKind == JsonValueKind.String) sb.Append(t.GetString());
                        }
                        content = sb.ToString();
                    }
                }
                if (!string.IsNullOrEmpty(content))
                {
                    // Extract JSON
                    try
                    {
                        var cellsDoc = JsonDocument.Parse(content);
                        if (cellsDoc.RootElement.ValueKind == JsonValueKind.Object && cellsDoc.RootElement.TryGetProperty("cells", out var cellsEl))
                        {
                            var rows = ExternalApiProvider.ParseCellsArray(cellsEl);
                            result.Cells = ExternalApiProvider.FitToShape(rows, context.Rows, context.Cols);
                            return result;
                        }
                    }
                    catch { }
                    // Fallback: split lines
                    var lines = content.Replace("\r", string.Empty).Split('\n');
                    result.Cells = ExternalApiProvider.FitToShape(lines, context.Rows, context.Cols);
                    return result;
                }
            }
            result.Cells = ExternalApiProvider.FitToShape(Array.Empty<string>(), context.Rows, context.Cols);
            return result;
        }
    }
}
