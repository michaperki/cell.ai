using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace SpreadsheetApp.Core.AI
{
    public sealed class ExternalApiProvider : IInferenceProvider
    {
        private readonly string _endpoint;
        private readonly string _apiKey;
        private readonly HttpClient _http;

        public ExternalApiProvider(string endpoint, string apiKey, HttpClient? http = null)
        {
            _endpoint = endpoint ?? throw new ArgumentNullException(nameof(endpoint));
            _apiKey = apiKey ?? throw new ArgumentNullException(nameof(apiKey));
            _http = http ?? new HttpClient();
            _http.Timeout = TimeSpan.FromSeconds(15);
        }

        public async Task<AIResult> GenerateFillAsync(AIContext context, CancellationToken cancellationToken)
        {
            var payload = new
            {
                sheetName = context.SheetName,
                startRow = context.StartRow,
                startCol = context.StartCol,
                rows = context.Rows,
                cols = context.Cols,
                title = context.Title,
                prompt = context.Prompt
            };
            using var req = new HttpRequestMessage(HttpMethod.Post, _endpoint);
            req.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            if (!string.IsNullOrEmpty(_apiKey)) req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);
            req.Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");

            using var resp = await _http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false);
            resp.EnsureSuccessStatusCode();
            await using var stream = await resp.Content.ReadAsStreamAsync(cancellationToken).ConfigureAwait(false);
            using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken).ConfigureAwait(false);
            if (doc.RootElement.ValueKind == JsonValueKind.Object && doc.RootElement.TryGetProperty("cells", out var cellsEl) && cellsEl.ValueKind == JsonValueKind.Array)
            {
                var list = ParseCellsArray(cellsEl);
                return new AIResult { Cells = FitToShape(list, context.Rows, context.Cols) };
            }
            // Fallback: look for text and split by lines
            string? text = null;
            if (doc.RootElement.TryGetProperty("text", out var textEl) && textEl.ValueKind == JsonValueKind.String)
                text = textEl.GetString();
            else if (doc.RootElement.TryGetProperty("content", out var contEl) && contEl.ValueKind == JsonValueKind.String)
                text = contEl.GetString();
            if (!string.IsNullOrEmpty(text))
            {
                var lines = text.Replace("\r", string.Empty).Split('\n');
                return new AIResult { Cells = FitToShape(lines, context.Rows, context.Cols) };
            }
            // As a last resort, empty payload with requested shape
            return new AIResult { Cells = FitToShape(Array.Empty<string>(), context.Rows, context.Cols) };
        }

        internal static string[][] ParseCellsArray(JsonElement cellsEl)
        {
            var rows = new System.Collections.Generic.List<string[]>();
            foreach (var rowEl in cellsEl.EnumerateArray())
            {
                var cols = new System.Collections.Generic.List<string>();
                if (rowEl.ValueKind == JsonValueKind.Array)
                {
                    foreach (var c in rowEl.EnumerateArray()) cols.Add(c.GetString() ?? string.Empty);
                }
                else
                {
                    cols.Add(rowEl.GetString() ?? string.Empty);
                }
                rows.Add(cols.ToArray());
            }
            return rows.ToArray();
        }

        internal static string[][] FitToShape(System.Collections.Generic.IEnumerable<string> items, int rows, int cols)
        {
            var list = new System.Collections.Generic.List<string>(items);
            return FitToShape(list, rows, cols);
        }

        internal static string[][] FitToShape(string[][] data, int rows, int cols)
        {
            var flat = new System.Collections.Generic.List<string>();
            foreach (var r in data)
            {
                foreach (var v in r) flat.Add(v ?? string.Empty);
            }
            return FitToShape(flat, rows, cols);
        }

        internal static string[][] FitToShape(System.Collections.Generic.List<string> flat, int rows, int cols)
        {
            rows = Math.Max(1, rows); cols = Math.Max(1, cols);
            var outArr = new string[rows][];
            int k = 0;
            for (int r = 0; r < rows; r++)
            {
                outArr[r] = new string[cols];
                for (int c = 0; c < cols; c++) outArr[r][c] = k < flat.Count ? flat[k++] : string.Empty;
            }
            return outArr;
        }
    }
}
