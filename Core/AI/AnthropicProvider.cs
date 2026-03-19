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
            var body = new
            {
                model = _model,
                max_tokens = 800,
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
            using var resp = await _http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false);
            resp.EnsureSuccessStatusCode();
            await using var stream = await resp.Content.ReadAsStreamAsync(cancellationToken).ConfigureAwait(false);
            using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken).ConfigureAwait(false);
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
                            return new AIResult { Cells = ExternalApiProvider.FitToShape(rows, context.Rows, context.Cols) };
                        }
                    }
                    catch { }
                    var lines = text.Replace("\r", string.Empty).Split('\n');
                    return new AIResult { Cells = ExternalApiProvider.FitToShape(lines, context.Rows, context.Cols) };
                }
            }
            return new AIResult { Cells = ExternalApiProvider.FitToShape(Array.Empty<string>(), context.Rows, context.Cols) };
        }
    }
}

