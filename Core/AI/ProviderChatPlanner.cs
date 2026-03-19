using System;
using System.Text.Json;
using System.Text;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;
using SpreadsheetApp.Core;

namespace SpreadsheetApp.Core.AI
{
    public sealed class ProviderChatPlanner : IChatPlanner
    {
        private readonly string _provider; // OpenAI, Anthropic, Auto

        public ProviderChatPlanner(string provider)
        {
            _provider = string.IsNullOrWhiteSpace(provider) ? "Auto" : provider;
        }

        public async Task<AIPlan> PlanAsync(AIContext context, string prompt, CancellationToken cancellationToken)
        {
            string provider = _provider;
            if (string.Equals(provider, "Auto", StringComparison.OrdinalIgnoreCase))
            {
                if (!string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("OPENAI_API_KEY"))) provider = "OpenAI";
                else if (!string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY"))) provider = "Anthropic";
                else return new MockChatPlanner().PlanAsync(context, prompt, cancellationToken).Result;
            }

            string sys = "You are a spreadsheet planning assistant. Respond ONLY with strict JSON matching this schema: {\"commands\":[{\"type\":\"set_values\",\"start\":{\"row\":<1-based int>,\"col\":<column letter>},\"values\":[[\"text\"],...]},{\"type\":\"set_title\",\"start\":{\"row\":<1-based>,\"col\":<letter>},\"rows\":1,\"cols\":1,\"text\":\"...\"},{\"type\":\"create_sheet\",\"name\":\"...\"},{\"type\":\"clear_range\",\"start\":{\"row\":<1-based>,\"col\":<letter>},\"rows\":<int>,\"cols\":<int>},{\"type\":\"rename_sheet\",\"index\":<1-based optional>,\"old_name\":\"... optional\",\"new_name\":\"...\"}]} with no extra keys, no prose.";
            string usr = $"Sheet={context.SheetName}; Selection=({context.StartRow+1},{CellAddress.ColumnIndexToName(context.StartCol)}); Rows={context.Rows}; Cols={context.Cols}; Title={(context.Title??string.Empty)}; Instruction={(prompt ?? string.Empty)}. Keep total writes <= 5000. Prefer list fills near the selection. Use set_values for rectangular fills.";

            string json = provider switch
            {
                "OpenAI" => await CallOpenAIAsync(sys, usr, cancellationToken).ConfigureAwait(false),
                "Anthropic" => await CallAnthropicAsync(sys, usr, cancellationToken).ConfigureAwait(false),
                _ => string.Empty
            };
            if (string.IsNullOrWhiteSpace(json)) return new AIPlan();
            try
            {
                var doc = JsonDocument.Parse(json);
                return ParsePlan(doc);
            }
            catch
            {
                // Try to extract JSON substring if model wrapped it
                int i1 = json.IndexOf('{');
                int i2 = json.LastIndexOf('}');
                if (i1 >= 0 && i2 > i1)
                {
                    try { var doc = JsonDocument.Parse(json.Substring(i1, i2 - i1 + 1)); return ParsePlan(doc); }
                    catch { }
                }
                return new AIPlan();
            }
        }

        private static async Task<string> CallOpenAIAsync(string system, string user, CancellationToken ct)
        {
            var key = Environment.GetEnvironmentVariable("OPENAI_API_KEY");
            if (string.IsNullOrWhiteSpace(key)) return string.Empty;
            var model = Environment.GetEnvironmentVariable("OPENAI_MODEL") ?? "gpt-4o-mini";
            var endpoint = Environment.GetEnvironmentVariable("OPENAI_BASE_URL") ?? "https://api.openai.com/v1/chat/completions";
            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(20) };
            http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", key);
            http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            var body = new
            {
                model,
                temperature = 0.2,
                messages = new object[] { new { role = "system", content = system }, new { role = "user", content = user } }
            };
            using var req = new HttpRequestMessage(HttpMethod.Post, endpoint)
            {
                Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json")
            };
            using var resp = await http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false);
            resp.EnsureSuccessStatusCode();
            await using var stream = await resp.Content.ReadAsStreamAsync(ct).ConfigureAwait(false);
            using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: ct).ConfigureAwait(false);
            if (doc.RootElement.TryGetProperty("choices", out var choices) && choices.ValueKind == JsonValueKind.Array && choices.GetArrayLength() > 0)
            {
                var msg = choices[0].GetProperty("message");
                if (msg.TryGetProperty("content", out var cnt))
                {
                    if (cnt.ValueKind == JsonValueKind.String) return cnt.GetString() ?? string.Empty;
                    if (cnt.ValueKind == JsonValueKind.Array)
                    {
                        var sb = new StringBuilder();
                        foreach (var part in cnt.EnumerateArray()) if (part.TryGetProperty("text", out var t) && t.ValueKind == JsonValueKind.String) sb.Append(t.GetString());
                        return sb.ToString();
                    }
                }
            }
            return string.Empty;
        }

        private static async Task<string> CallAnthropicAsync(string system, string user, CancellationToken ct)
        {
            var key = Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY");
            if (string.IsNullOrWhiteSpace(key)) return string.Empty;
            var model = Environment.GetEnvironmentVariable("ANTHROPIC_MODEL") ?? "claude-3-haiku-20240307";
            var endpoint = Environment.GetEnvironmentVariable("ANTHROPIC_BASE_URL") ?? "https://api.anthropic.com/v1/messages";
            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(20) };
            http.DefaultRequestHeaders.Add("x-api-key", key);
            http.DefaultRequestHeaders.Add("anthropic-version", "2023-06-01");
            var body = new
            {
                model,
                max_tokens = 800,
                temperature = 0.2,
                system = system,
                messages = new object[] { new { role = "user", content = new object[] { new { type = "text", text = user } } } }
            };
            using var req = new HttpRequestMessage(HttpMethod.Post, endpoint)
            {
                Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json")
            };
            using var resp = await http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false);
            resp.EnsureSuccessStatusCode();
            await using var stream = await resp.Content.ReadAsStreamAsync(ct).ConfigureAwait(false);
            using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: ct).ConfigureAwait(false);
            if (doc.RootElement.TryGetProperty("content", out var contentEl) && contentEl.ValueKind == JsonValueKind.Array)
            {
                foreach (var part in contentEl.EnumerateArray())
                {
                    if (part.TryGetProperty("text", out var t) && t.ValueKind == JsonValueKind.String) return t.GetString() ?? string.Empty;
                }
            }
            return string.Empty;
        }

        private static AIPlan ParsePlan(JsonDocument doc)
        {
            var plan = new AIPlan();
            if (doc.RootElement.ValueKind != JsonValueKind.Object) return plan;
            if (!doc.RootElement.TryGetProperty("commands", out var cmds) || cmds.ValueKind != JsonValueKind.Array) return plan;
            foreach (var cmd in cmds.EnumerateArray())
            {
                try
                {
                    if (!cmd.TryGetProperty("type", out var t)) continue;
                    var type = t.GetString()?.Trim().ToLowerInvariant();
                    switch (type)
                    {
                        case "set_values":
                            {
                                var sv = new SetValuesCommand();
                                if (cmd.TryGetProperty("start", out var start))
                                {
                                    int srr, scc;
                                    ParseStart(start, out srr, out scc);
                                    sv.StartRow = srr;
                                    sv.StartCol = scc;
                                }
                                if (cmd.TryGetProperty("values", out var vals) && vals.ValueKind == JsonValueKind.Array)
                                {
                                    var rows = new System.Collections.Generic.List<string[]>();
                                    foreach (var rowEl in vals.EnumerateArray())
                                    {
                                        var cols = new System.Collections.Generic.List<string>();
                                        if (rowEl.ValueKind == JsonValueKind.Array)
                                            foreach (var c in rowEl.EnumerateArray()) cols.Add(c.GetString() ?? string.Empty);
                                        rows.Add(cols.ToArray());
                                    }
                                    sv.Values = rows.ToArray();
                                }
                                plan.Commands.Add(sv);
                            break;
                        }
                        case "clear_range":
                            {
                                var cr = new ClearRangeCommand();
                                if (cmd.TryGetProperty("start", out var start))
                                {
                                    int srr, scc;
                                    ParseStart(start, out srr, out scc);
                                    cr.StartRow = srr; cr.StartCol = scc;
                                }
                                if (cmd.TryGetProperty("rows", out var r)) cr.Rows = SafeInt(r, 1);
                                if (cmd.TryGetProperty("cols", out var c)) cr.Cols = SafeInt(c, 1);
                                plan.Commands.Add(cr);
                                break;
                            }
                        case "set_title":
                            {
                                var st = new SetTitleCommand();
                                if (cmd.TryGetProperty("start", out var start))
                                {
                                    int srr, scc;
                                    ParseStart(start, out srr, out scc);
                                    st.StartRow = srr; st.StartCol = scc;
                                }
                                if (cmd.TryGetProperty("rows", out var r)) st.Rows = SafeInt(r, 1);
                                if (cmd.TryGetProperty("cols", out var c)) st.Cols = SafeInt(c, 1);
                                if (cmd.TryGetProperty("text", out var tx)) st.Text = tx.GetString() ?? string.Empty;
                                plan.Commands.Add(st);
                                break;
                            }
                        case "rename_sheet":
                            {
                                var rn = new RenameSheetCommand();
                                if (cmd.TryGetProperty("index", out var idx)) rn.Index1 = SafeInt(idx, 0) > 0 ? SafeInt(idx, 0) : (int?)null;
                                if (cmd.TryGetProperty("old_name", out var on) && on.ValueKind == JsonValueKind.String) rn.OldName = on.GetString();
                                if (cmd.TryGetProperty("new_name", out var nn) && nn.ValueKind == JsonValueKind.String) rn.NewName = nn.GetString() ?? rn.NewName;
                                plan.Commands.Add(rn);
                                break;
                            }
                        case "create_sheet":
                            {
                                var cs = new CreateSheetCommand();
                                if (cmd.TryGetProperty("name", out var n)) cs.Name = n.GetString() ?? cs.Name;
                                plan.Commands.Add(cs);
                                break;
                            }
                    }
                }
                catch { }
            }
            return plan;
        }

        private static void ParseStart(JsonElement start, out int rowZero, out int colZero)
        {
            int row1 = 1;
            string? colLetter = "A";
            if (start.ValueKind == JsonValueKind.Object)
            {
                if (start.TryGetProperty("row", out var r)) row1 = SafeInt(r, 1);
                if (start.TryGetProperty("col", out var c))
                {
                    if (c.ValueKind == JsonValueKind.String) colLetter = c.GetString();
                    else colLetter = CellAddress.ColumnIndexToName(SafeInt(c, 1) - 1);
                }
            }
            rowZero = Math.Max(0, row1 - 1);
            try { colZero = CellAddress.ColumnNameToIndex(colLetter ?? "A"); }
            catch { colZero = 0; }
        }

        private static int SafeInt(JsonElement el, int def)
        {
            try { return el.GetInt32(); } catch { try { return (int)el.GetDouble(); } catch { return def; } }
        }
    }
}
