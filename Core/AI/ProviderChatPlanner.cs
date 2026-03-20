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
                else
                {
                    var mockPlan = new MockChatPlanner().PlanAsync(context, prompt, cancellationToken).Result;
                    mockPlan.RawJson = string.Empty;
                    return mockPlan;
                }
            }

            string sys = "You are a spreadsheet planning assistant. Respond ONLY with strict JSON matching this schema: {\"commands\":[{\"type\":\"set_values\",\"start\":{\"row\":<1-based int>,\"col\":<column letter>},\"values\":[[\"text\"],...]},{\"type\":\"set_formula\",\"start\":{\"row\":<1-based>,\"col\":<letter>},\"formulas\":[[\"=A1+B1\"],...]},{\"type\":\"set_title\",\"start\":{\"row\":<1-based>,\"col\":<letter>},\"rows\":1,\"cols\":1,\"text\":\"...\"},{\"type\":\"create_sheet\",\"name\":\"...\"},{\"type\":\"clear_range\",\"start\":{\"row\":<1-based>,\"col\":<letter>},\"rows\":<int>,\"cols\":<int>},{\"type\":\"rename_sheet\",\"index\":<1-based optional>,\"old_name\":\"... optional\",\"new_name\":\"...\"},{\"type\":\"sort_range\",\"start\":{\"row\":<1-based>,\"col\":<letter>},\"rows\":<int>,\"cols\":<int>,\"sort_col\":\"<letter or 1-based index>\",\"order\":\"asc|desc\",\"has_header\":<bool> }]} with no extra keys, no prose. Only perform the requested change(s). Do NOT add titles, totals, or extra columns unless explicitly asked. When creating tables, place headers at the start cell's row and write data rows immediately below. If a selection/range shape is indicated (Rows/Cols), align your writes to that shape and avoid writing outside it. When filling a table from a list of inputs, combine all rows into a single set_values command with a 2D values array.";

            // Allowed commands and strengthened constraints
            string[]? allowedCmds = null;
            try { allowedCmds = (context.AllowedCommands != null && context.AllowedCommands.Length > 0) ? context.AllowedCommands : null; } catch { allowedCmds = null; }
            // Back-compat: infer valuesOnly/noTitles from prompt if no explicit policy was provided
            bool inferredValuesOnly = (prompt ?? string.Empty).IndexOf("set_values only", StringComparison.OrdinalIgnoreCase) >= 0;
            bool inferredNoTitles = inferredValuesOnly || (prompt ?? string.Empty).IndexOf("do not add title", StringComparison.OrdinalIgnoreCase) >= 0 || (prompt ?? string.Empty).IndexOf("do not add titles", StringComparison.OrdinalIgnoreCase) >= 0;
            if (allowedCmds != null)
            {
                // Emit an explicit AllowedCommands list into the system message
                var sbAllowed = new System.Text.StringBuilder();
                sbAllowed.Append(" Allowed command types: ");
                for (int i = 0; i < allowedCmds.Length; i++)
                {
                    if (i > 0) sbAllowed.Append(", ");
                    sbAllowed.Append(allowedCmds[i]);
                }
                sbAllowed.Append('.')
                    .Append(" Use only these commands and no others.");
                sys += sbAllowed.ToString();
            }
            else if (inferredValuesOnly)
            {
                sys += " Allowed command types: set_values only. Do NOT use set_title or set_formula or any other commands.";
            }
            else if (inferredNoTitles)
            {
                sys += " Do NOT use set_title for this request.";
            }

            var sbUsr = new System.Text.StringBuilder();
            sbUsr.Append($"Sheet={context.SheetName}; Selection=({context.StartRow+1},{CellAddress.ColumnIndexToName(context.StartCol)}); Rows={context.Rows}; Cols={context.Cols}; Title={(context.Title??string.Empty)}. ");
            // Include small snapshots to improve planning
            try
            {
                if (context.SelectionValues != null && context.SelectionValues.Length > 0)
                {
                    sbUsr.Append(" SelectionContent=");
                    int rlim = Math.Min(context.SelectionValues.Length, 20);
                    for (int i = 0; i < rlim; i++)
                    {
                        if (i > 0) sbUsr.Append(" || ");
                        var row = context.SelectionValues[i];
                        int clim = Math.Min(row.Length, 10);
                        sbUsr.Append(string.Join(" | ", row.AsSpan(0, clim).ToArray()));
                    }
                    sbUsr.Append(";");
                }
                if (context.NearbyValues != null && context.NearbyValues.Length > 0)
                {
                    sbUsr.Append(" Nearby=");
                    int rlim = Math.Min(context.NearbyValues.Length, 20);
                    for (int i = 0; i < rlim; i++)
                    {
                        if (i > 0) sbUsr.Append(" || ");
                        var row = context.NearbyValues[i];
                        int clim = Math.Min(row.Length, 10);
                        sbUsr.Append(string.Join(" | ", row.AsSpan(0, clim).ToArray()));
                    }
                    sbUsr.Append(";");
                }
                if (context.Workbook != null && context.Workbook.Length > 0)
                {
                    sbUsr.Append(" Workbook=");
                    int lim = Math.Min(context.Workbook.Length, 10);
                    for (int i = 0; i < lim; i++)
                    {
                        var s = context.Workbook[i];
                        if (i > 0) sbUsr.Append("; ");
                        string hdr = (s.HeaderRow != null ? string.Join(",", s.HeaderRow) : string.Empty);
                        string used = (!string.IsNullOrWhiteSpace(s.UsedTopLeft) && !string.IsNullOrWhiteSpace(s.UsedBottomRight)) ? ($"{s.UsedTopLeft}:{s.UsedBottomRight}") : string.Empty;
                        sbUsr.Append($"[{s.Name} rows={s.UsedRows} cols={s.UsedCols} header_idx={s.HeaderRowIndex} data_rows={s.DataRowCountExcludingHeader} used={used} header={hdr}]");
                    }
                    sbUsr.Append(". ");
                }
                if (context.Conversation != null && context.Conversation.Count > 0)
                {
                    sbUsr.Append(" History=");
                    int start = Math.Max(0, context.Conversation.Count - 6);
                    for (int i = start; i < context.Conversation.Count; i++)
                    {
                        var m = context.Conversation[i];
                        sbUsr.Append($"[{m.Role}:{m.Content}]");
                    }
                    sbUsr.Append(". ");
                }
            }
            catch { }
            var schemaSection = TryBuildSchemaFillSection(context);
            if (!string.IsNullOrEmpty(schemaSection)) sbUsr.Append(schemaSection);
            sbUsr.Append($" Instruction={(prompt ?? string.Empty)}. Keep total writes <= 5000. Prefer list fills near the selection. Use set_values for plain text and set_formula for formulas; use sort_range for sorting.");
            string usr = sbUsr.ToString();

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
                var plan = ParsePlan(doc);
                // Filter disallowed command types: explicit AllowedCommands first, then inferred guards
                if (allowedCmds != null && allowedCmds.Length > 0)
                {
                    plan.Commands.RemoveAll(c => !IsCommandAllowedByList(c, allowedCmds));
                }
                else if (inferredValuesOnly)
                {
                    plan.Commands.RemoveAll(c => c is not SetValuesCommand);
                }
                else if (inferredNoTitles)
                {
                    plan.Commands.RemoveAll(c => c is SetTitleCommand);
                }
                plan.RawJson = json;
                plan.RawUser = usr;
                plan.RawSystem = sys;
                return plan;
            }
            catch
            {
                // Try to extract JSON substring if model wrapped it
                int i1 = json.IndexOf('{');
                int i2 = json.LastIndexOf('}');
                if (i1 >= 0 && i2 > i1)
                {
                    try { var s = json.Substring(i1, i2 - i1 + 1); var doc = JsonDocument.Parse(s); var plan = ParsePlan(doc); plan.RawJson = s; plan.RawUser = usr; plan.RawSystem = sys; return plan; }
                    catch { }
                }
                return new AIPlan();
            }
        }

        private static string? TryBuildSchemaFillSection(AIContext context)
        {
            // Prefer workbook header summary for robust schema, fall back to Nearby header row
            string[]? headerRow = null;
            int tableLeftCol = 0; // absolute column index of the table's first column (e.g., A=0)
            try
            {
                if (context.Workbook != null)
                {
                    foreach (var s in context.Workbook)
                    {
                        if (string.Equals(s.Name, context.SheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            headerRow = s.HeaderRow;
                            if (!string.IsNullOrWhiteSpace(s.UsedTopLeft) && CellAddress.TryParse(s.UsedTopLeft, out int _, out int col))
                                tableLeftCol = Math.Max(0, col);
                            break;
                        }
                    }
                }
            }
            catch { }

            var nearby = context.NearbyValues;
            if (headerRow == null || headerRow.Length < 2)
            {
                if (nearby == null || nearby.Length < 2) return null;
                headerRow = nearby[0];
            }

            // Require at least two non-empty headers beyond the input column
            int nonEmptyHeaders = 0;
            for (int i = 1; i < headerRow.Length; i++) if (!string.IsNullOrWhiteSpace(headerRow[i])) nonEmptyHeaders++;
            if (nonEmptyHeaders < 2) return null;

            // Build the section
            var sb = new System.Text.StringBuilder();
            sb.Append(" FillMapping (emit ONE set_values with a 2D array):");

            // Headers — use absolute letters from the table's left edge; skip input col (index 0)
            sb.Append(" Headers:");
            for (int c = 1; c < headerRow.Length; c++)
            {
                if (string.IsNullOrWhiteSpace(headerRow[c])) continue;
                string colLetter = CellAddress.ColumnIndexToName(tableLeftCol + c);
                if (c > 1) sb.Append(',');
                sb.Append($" {colLetter}={headerRow[c]}");
            }
            sb.Append('.');

            // Policy: input column and writable columns
            try
            {
                var policy = context.WritePolicy;
                if (policy != null)
                {
                    if (policy.WritableColumns != null && policy.WritableColumns.Length > 0)
                    {
                        sb.Append(" WritableColumns=");
                        for (int i = 0; i < policy.WritableColumns.Length; i++)
                        {
                            if (i > 0) sb.Append(',');
                            sb.Append(CellAddress.ColumnIndexToName(policy.WritableColumns[i]));
                        }
                        sb.Append('.');
                    }
                    if (policy.InputColumnIndex.HasValue)
                    {
                        string inputLetter = CellAddress.ColumnIndexToName(Math.Max(0, policy.InputColumnIndex.Value));
                        if (!policy.AllowInputWritesForExistingRows && !policy.AllowInputWritesForEmptyRows)
                        {
                            sb.Append($" InputColumn={inputLetter}. Do not write to {inputLetter}.");
                        }
                        else if (!policy.AllowInputWritesForExistingRows && policy.AllowInputWritesForEmptyRows)
                        {
                            sb.Append($" InputColumn={inputLetter}. Do not write to existing values in {inputLetter}; writing to empty rows within the selection is allowed.");
                        }
                        else if (policy.AllowInputWritesForExistingRows && !policy.AllowInputWritesForEmptyRows)
                        {
                            sb.Append($" InputColumn={inputLetter}. Only existing rows in {inputLetter} may be edited; do not write new inputs.");
                        }
                        else
                        {
                            sb.Append($" InputColumn={inputLetter}. Writing is allowed.");
                        }
                    }
                }
                else
                {
                    // Default: do not write to the first (input) column
                    string inputColLetter = CellAddress.ColumnIndexToName(tableLeftCol + 0);
                    sb.Append($" InputColumn={inputColLetter}. Do not write to {inputColLetter}; write outputs only to the mapped columns.");
                }
            }
            catch { }

            // Row-to-input mapping: list inputs detected in the first column of the Nearby window
            if (nearby != null && nearby.Length > 1)
            {
                int startRow1 = context.StartRow + 1; // 1-based starting row for the selection
                int fillColCount = Math.Max(0, headerRow.Length - 1);
                string fillStart = CellAddress.ColumnIndexToName(tableLeftCol + 1);
                string fillEnd = CellAddress.ColumnIndexToName(tableLeftCol + fillColCount);
                for (int r = 1; r < nearby.Length; r++)
                {
                    if (nearby[r].Length > 0 && !string.IsNullOrWhiteSpace(nearby[r][0]))
                    {
                        sb.Append($" Row {startRow1 + r - 1}: input=\"{nearby[r][0]}\" -> fill columns {fillStart}-{fillEnd}.");
                    }
                }
            }

            // Schema: include column headers for the selection if provided
            try
            {
                if (context.Schema != null && context.Schema.Length > 0)
                {
                    sb.Append(" Schema:");
                    int count = 0;
                    foreach (var col in context.Schema)
                    {
                        if (count > 0) sb.Append(',');
                        sb.Append(' ').Append(col.ColumnLetter);
                        if (!string.IsNullOrWhiteSpace(col.Name)) sb.Append('=').Append(col.Name);
                        sb.Append(" (").Append(col.Type)
                          .Append(col.AllowEmpty ? ", optional)" : ", required)");
                        count++;
                    }
                    sb.Append('.');
                }
            }
            catch { }

            return sb.ToString();
        }

        private static bool IsCommandAllowedByList(IAICommand cmd, string[] allowed)
        {
            string type = cmd switch
            {
                SetValuesCommand => "set_values",
                SetFormulaCommand => "set_formula",
                SetTitleCommand => "set_title",
                CreateSheetCommand => "create_sheet",
                ClearRangeCommand => "clear_range",
                RenameSheetCommand => "rename_sheet",
                SortRangeCommand => "sort_range",
                _ => string.Empty
            };
            if (string.IsNullOrEmpty(type)) return false;
            foreach (var a in allowed)
            {
                if (string.Equals(a?.Trim(), type, StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
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
                temperature = 0.0,
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
            int maxTokens = 2048;
            try { var s = Environment.GetEnvironmentVariable("ANTHROPIC_MAX_TOKENS"); if (!string.IsNullOrWhiteSpace(s)) maxTokens = int.Parse(s); } catch { }
            var body = new
            {
                model,
                max_tokens = maxTokens,
                temperature = 0.0,
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
                                        {
                                            foreach (var ce in rowEl.EnumerateArray())
                                            {
                                                string cellText = string.Empty;
                                                switch (ce.ValueKind)
                                                {
                                                    case JsonValueKind.String:
                                                        cellText = ce.GetString() ?? string.Empty; break;
                                                    case JsonValueKind.Number:
                                                    case JsonValueKind.Object:
                                                    case JsonValueKind.Array:
                                                        cellText = ce.GetRawText(); break;
                                                    case JsonValueKind.True:
                                                        cellText = "true"; break;
                                                    case JsonValueKind.False:
                                                        cellText = "false"; break;
                                                    case JsonValueKind.Null:
                                                    default:
                                                        cellText = string.Empty; break;
                                                }
                                                cols.Add(cellText);
                                            }
                                        }
                                        rows.Add(cols.ToArray());
                                    }
                                    sv.Values = rows.ToArray();
                                }
                                plan.Commands.Add(sv);
                            break;
                        }
                        case "set_formula":
                            {
                                var sf = new SetFormulaCommand();
                                if (cmd.TryGetProperty("start", out var start))
                                {
                                    int srr, scc; ParseStart(start, out srr, out scc); sf.StartRow = srr; sf.StartCol = scc;
                                }
                                if (cmd.TryGetProperty("formulas", out var vals) && vals.ValueKind == JsonValueKind.Array)
                                {
                                    var rows = new System.Collections.Generic.List<string[]>();
                                    foreach (var rowEl in vals.EnumerateArray())
                                    {
                                        var cols = new System.Collections.Generic.List<string>();
                                        if (rowEl.ValueKind == JsonValueKind.Array)
                                        {
                                            foreach (var ce in rowEl.EnumerateArray())
                                            {
                                                string textVal = ce.ValueKind == JsonValueKind.String ? (ce.GetString() ?? string.Empty) : ce.GetRawText();
                                                cols.Add(textVal);
                                            }
                                        }
                                        rows.Add(cols.ToArray());
                                    }
                                    sf.Formulas = rows.ToArray();
                                }
                                plan.Commands.Add(sf);
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
                        case "sort_range":
                            {
                                var sr = new SortRangeCommand();
                                if (cmd.TryGetProperty("start", out var start))
                                {
                                    int srr, scc; ParseStart(start, out srr, out scc); sr.StartRow = srr; sr.StartCol = scc;
                                }
                                if (cmd.TryGetProperty("rows", out var rr)) sr.Rows = SafeInt(rr, 1);
                                if (cmd.TryGetProperty("cols", out var cc)) sr.Cols = SafeInt(cc, 1);
                                if (cmd.TryGetProperty("order", out var ord) && ord.ValueKind == JsonValueKind.String) sr.Order = (ord.GetString() ?? "asc").ToLowerInvariant();
                                if (cmd.TryGetProperty("has_header", out var hh) && (hh.ValueKind == JsonValueKind.True || hh.ValueKind == JsonValueKind.False)) sr.HasHeader = hh.GetBoolean();
                                // sort_col can be string letter (absolute) or number (1-based relative)
                                if (cmd.TryGetProperty("sort_col", out var sc))
                                {
                                    if (sc.ValueKind == JsonValueKind.String)
                                    {
                                        string? letter = sc.GetString();
                                        try { sr.SortCol = CellAddress.ColumnNameToIndex(letter ?? "A"); }
                                        catch { sr.SortCol = sr.StartCol; }
                                    }
                                    else
                                    {
                                        int idx1 = SafeInt(sc, 1);
                                        sr.SortCol = Math.Max(0, sr.StartCol + idx1 - 1);
                                    }
                                }
                                plan.Commands.Add(sr);
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
