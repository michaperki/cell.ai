using System;
using System.Text.Json;
using System.Text;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;
using SpreadsheetApp.Core;
using System.IO;

namespace SpreadsheetApp.Core.AI
{
    public sealed class ProviderChatPlanner : IChatPlanner
    {
        private readonly string _provider; // OpenAI, Anthropic, Auto
        public static int? ActivePromptVersion { get; private set; }
        public static string? ActivePromptPath { get; private set; }

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

            string sys;
            if (context.RequestQueriesOnly)
            {
                if (context.AnswerOnly)
                {
                    // Ask mode: return a direct textual answer as JSON {"answer":"..."}
                    sys = "You are a spreadsheet Q&A assistant. Respond ONLY with strict JSON: {\"answer\":\"<short answer>\"}. Do not include any other keys, no code blocks, no prose outside JSON. Use the provided selection and headers as context to answer the user's question succinctly.";
                }
                else
                {
                    // Observation mode for agent loop: return queries
                    sys = "You are a spreadsheet planning assistant. Respond ONLY with strict JSON matching this schema: {\"queries\":[{\"type\":\"selection_summary\"},{\"type\":\"profile_column\",\"col\":\"<letter or 1-based index>\",\"rows\":<int optional>},{\"type\":\"describe_column\",\"col\":\"<letter or 1-based index>\",\"rows\":<int optional>},{\"type\":\"unique_values\",\"col\":\"<letter or 1-based index>\",\"top\":<int optional>},{\"type\":\"sample_rows\",\"rows\":<int>,\"cols\":<int>},{\"type\":\"count_where\",\"filters\":[{\"col\":\"<letter or 1-based index>\",\"op\":\"eq|ne|gt|ge|lt|le|contains|not_contains\",\"value\":\"...\"}]}]} with no extra keys, no prose. Choose a small set of high-value queries to understand the selection (uniques, column descriptions, a few sample rows, and counts under simple filters). Do not include write commands.";
                }
            }
            else
            {
                sys = "You are a spreadsheet planning assistant. Respond ONLY with strict JSON matching this schema: {\"commands\":[{\"type\":\"set_values\",\"start\":{\"row\":<1-based int>,\"col\":<column letter>},\"values\":[[\"text\"],...]},{\"type\":\"set_formula\",\"start\":{\"row\":<1-based>,\"col\":<letter>},\"formulas\":[[\"=A1+B1\"],...]},{\"type\":\"set_title\",\"start\":{\"row\":<1-based>,\"col\":<letter>},\"rows\":1,\"cols\":1,\"text\":\"...\"},{\"type\":\"create_sheet\",\"name\":\"...\"},{\"type\":\"clear_range\",\"start\":{\"row\":<1-based>,\"col\":<letter>},\"rows\":<int>,\"cols\":<int>},{\"type\":\"rename_sheet\",\"index\":<1-based optional>,\"old_name\":\"... optional\",\"new_name\":\"...\"},{\"type\":\"sort_range\",\"start\":{\"row\":<1-based>,\"col\":<letter>},\"rows\":<int>,\"cols\":<int>,\"sort_col\":\"<letter or 1-based index>\",\"order\":\"asc|desc\",\"has_header\":<bool> },{\"type\":\"insert_rows\",\"at\":<1-based row>,\"count\":<int>},{\"type\":\"delete_rows\",\"at\":<1-based row>,\"count\":<int>},{\"type\":\"insert_cols\",\"at\":\"<letter or 1-based index>\",\"count\":<int>},{\"type\":\"delete_cols\",\"at\":\"<letter or 1-based index>\",\"count\":<int>},{\"type\":\"delete_sheet\",\"index\":<1-based optional>,\"name\":\"... optional\"},{\"type\":\"copy_range\",\"start\":{\"row\":<1-based>,\"col\":<letter>},\"rows\":<int>,\"cols\":<int>,\"dest\":{\"row\":<1-based>,\"col\":<letter>}},{\"type\":\"move_range\",\"start\":{\"row\":<1-based>,\"col\":<letter>},\"rows\":<int>,\"cols\":<int>,\"dest\":{\"row\":<1-based>,\"col\":<letter>}},{\"type\":\"set_format\",\"start\":{\"row\":<1-based>,\"col\":<letter>},\"rows\":<int>,\"cols\":<int>,\"bold\":<bool optional>,\"number_format\":\"... optional\",\"h_align\":\"left|center|right optional\",\"fore_color\":\"#RRGGBB optional\",\"back_color\":\"#RRGGBB optional\"},{\"type\":\"set_validation\",\"start\":{\"row\":<1-based>,\"col\":<letter>},\"rows\":<int>,\"cols\":<int>,\"mode\":\"list|number_between\",\"allow_empty\":<bool>,\"min\":<number optional>,\"max\":<number optional>,\"allowed\":[\"a\",\"b\"]},{\"type\":\"set_conditional_format\",\"start\":{\"row\":<1-based>,\"col\":<letter>},\"rows\":<int>,\"cols\":<int>,\"op\":\">|>=|<|<=|==|!=\",\"threshold\":<number>,\"bold\":<bool optional>,\"number_format\":\"... optional\",\"h_align\":\"left|center|right optional\",\"fore_color\":\"#RRGGBB optional\",\"back_color\":\"#RRGGBB optional\"}]} with no extra keys, no prose. Only perform the requested change(s). Do NOT add titles, totals, or extra columns unless explicitly asked. When creating tables, place headers at the start cell's row and write data rows immediately below. If a selection/range shape is indicated (Rows/Cols), align your writes to that shape and avoid writing outside it. For two-column summary/label patterns, each set_values row MUST be exactly two columns [label, \"\"], and formulas belong in the second column via set_formula. When filling a table from a list of inputs, combine all rows into a single set_values command with a 2D values array. Normalize shorthand like '95k' to plain integers (e.g., 95000) when computing numeric values from text; always write plain numbers with no suffixes. Each command MAY include an optional \"rationale\" string explaining why; this is for preview only and has no effect on execution.";
                // Override with file-based system prompt if provided
                try
                {
                    var sysFromFile = TryLoadSystemPromptFromFile();
                    if (!string.IsNullOrWhiteSpace(sysFromFile)) sys = sysFromFile!;
                }
                catch { }
            }

            // Allowed commands and strengthened constraints
            string[]? allowedCmds = null;
            try { allowedCmds = (context.AllowedCommands != null && context.AllowedCommands.Length > 0) ? context.AllowedCommands : null; } catch { allowedCmds = null; }
            // Back-compat: infer valuesOnly/noTitles from prompt if no explicit policy was provided
            bool inferredValuesOnly = (prompt ?? string.Empty).IndexOf("set_values only", StringComparison.OrdinalIgnoreCase) >= 0;
            bool inferredNoTitles = inferredValuesOnly || (prompt ?? string.Empty).IndexOf("do not add title", StringComparison.OrdinalIgnoreCase) >= 0 || (prompt ?? string.Empty).IndexOf("do not add titles", StringComparison.OrdinalIgnoreCase) >= 0;

            // Structural-op strict gating: if caller didn't supply AllowedCommands, infer a minimal set from the prompt.
            if (allowedCmds == null)
            {
                var inferredAllowed = InferAllowedCommandsFromPrompt(prompt ?? string.Empty);
                if (inferredAllowed != null && inferredAllowed.Length > 0)
                {
                    allowedCmds = inferredAllowed;
                    // When we strictly gate to structural ops, ensure valuesOnly mode does not bias toward set_values
                    inferredValuesOnly = false;
                }
            }
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
            // Include FillMapping only for schema/table fills (pure set_values intent) and only when not in query mode
            bool includeFillMapping = false;
            try
            {
                if (!context.RequestQueriesOnly)
                {
                    if (allowedCmds != null && allowedCmds.Length == 1 && string.Equals(allowedCmds[0], "set_values", StringComparison.OrdinalIgnoreCase)) includeFillMapping = true;
                    else if (allowedCmds == null && inferredValuesOnly) includeFillMapping = true;
                }
            }
            catch { }
            if (includeFillMapping)
            {
                var schemaSection = TryBuildSchemaFillSection(context);
                if (!string.IsNullOrEmpty(schemaSection)) sbUsr.Append(schemaSection);
            }

            // Tailor guidance: if explicit AllowedCommands provided, avoid generic hints that bias toward set_values
            string guidance;
            bool valuesOnlyMode = (allowedCmds != null && allowedCmds.Length == 1 && string.Equals(allowedCmds[0], "set_values", StringComparison.OrdinalIgnoreCase)) || inferredValuesOnly;
            if (context.RequestQueriesOnly)
            {
                guidance = " Return only a small set of queries to better understand the data; avoid any write commands.";
            }
            else if (allowedCmds != null && allowedCmds.Length > 0)
            {
                var sb = new System.Text.StringBuilder();
                sb.Append(" Keep total writes <= 5000. Use only these commands: ");
                for (int i = 0; i < allowedCmds.Length; i++) { if (i > 0) sb.Append(',').Append(' '); sb.Append(allowedCmds[i]); }
                sb.Append(". Do not use any other commands.");
                guidance = sb.ToString();
            }
            else if (valuesOnlyMode)
            {
                guidance = " Keep total writes <= 5000. Prefer list fills near the selection. When formulas are needed, write them as strings beginning with '=' inside set_values cells so they evaluate as formulas. Do not use set_title or set_formula.";
            }
            else
            {
                guidance = " Keep total writes <= 5000. Use set_values for plain text and set_formula for formulas when needed; use sort_range for sorting.";
            }
            sbUsr.Append($" Instruction={(prompt ?? string.Empty)}.{guidance}");
            string usr = sbUsr.ToString();

            ChatCallResult call;
            try
            {
                call = provider switch
                {
                    "OpenAI" => await CallOpenAIWithUsageAsync(sys, usr, cancellationToken).ConfigureAwait(false),
                    "Anthropic" => await CallAnthropicWithUsageAsync(sys, usr, cancellationToken).ConfigureAwait(false),
                    _ => new ChatCallResult { Content = string.Empty, Provider = provider }
                };
            }
            catch (Exception ex)
            {
                // Network/provider failed — gracefully fall back to mock planner
                try { Console.Error.WriteLine($"[ProviderChatPlanner] API call failed, falling back to mock: {ex.GetType().Name}: {ex.Message}"); } catch { }
                try
                {
                    var mockPlan = new MockChatPlanner().PlanAsync(context, prompt, cancellationToken).Result;
                    mockPlan.RawJson = string.Empty; mockPlan.RawUser = usr; mockPlan.RawSystem = sys;
                    // Apply the same filtering/guards as provider-backed plans
                    if (allowedCmds != null && allowedCmds.Length > 0)
                    {
                        mockPlan.Commands.RemoveAll(c => !IsCommandAllowedByList(c, allowedCmds));
                    }
                    else if (inferredValuesOnly)
                    {
                        mockPlan.Commands.RemoveAll(c => c is not SetValuesCommand);
                    }
                    else if (inferredNoTitles)
                    {
                        mockPlan.Commands.RemoveAll(c => c is SetTitleCommand);
                    }
                    if (context.SelectionHardMode)
                    {
                        ApplySelectionHardMode(context, mockPlan);
                    }
                    return mockPlan;
                }
                catch { return new AIPlan(); }
            }
            string json = call.Content ?? string.Empty;
            if (string.IsNullOrWhiteSpace(json)) return new AIPlan();
            try
            {
                var doc = JsonDocument.Parse(json);
                var plan = ParsePlan(doc);
                // Attach provider usage metadata
                plan.Provider = string.IsNullOrWhiteSpace(call.Provider) ? provider : call.Provider;
                plan.Model = call.Model;
                plan.Usage = call.Usage;
                plan.LatencyMs = call.LatencyMs;
                // If this was query-only mode, return immediately after parsing queries
                if (context.RequestQueriesOnly)
                {
                    // Also capture Ask mode answers if present
                    try
                    {
                        if (doc.RootElement.ValueKind == JsonValueKind.Object && doc.RootElement.TryGetProperty("answer", out var ans) && ans.ValueKind == JsonValueKind.String)
                        {
                            plan.Answer = ans.GetString();
                        }
                    }
                    catch { }
                    plan.RawJson = json; plan.RawUser = usr; plan.RawSystem = sys;
                    return plan;
                }
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

                // Validate against selection bounds, schema width, and write policy. If violations exist, request one revision.
                var violations = ValidatePlanAgainstContext(context, plan, allowedCmds, inferredValuesOnly, inferredNoTitles, out int expectedWidth);
                if (violations.Count > 0)
                {
                    // Build a concise revision user message including constraints and first few violations
                    var rev = BuildRevisionUserMessage(context, usr, allowedCmds, expectedWidth, violations, plan.RawJson ?? string.Empty);
                    ChatCallResult call2;
                    try
                    {
                        call2 = provider switch
                        {
                            "OpenAI" => await CallOpenAIWithUsageAsync(sys, rev, cancellationToken).ConfigureAwait(false),
                            "Anthropic" => await CallAnthropicWithUsageAsync(sys, rev, cancellationToken).ConfigureAwait(false),
                            _ => new ChatCallResult { Content = string.Empty, Provider = provider }
                        };
                    }
                    catch (Exception ex2)
                    {
                        try { Console.Error.WriteLine($"[ProviderChatPlanner] Revision API call failed, falling back to mock: {ex2.GetType().Name}: {ex2.Message}"); } catch { }
                        try
                        {
                            var fallback = new MockChatPlanner().PlanAsync(context, prompt, cancellationToken).Result;
                            fallback.RawJson = string.Empty; fallback.RawUser = usr; fallback.RawSystem = sys;
                            // Apply allowed-commands filtering and selection fencing on fallback
                            if (allowedCmds != null && allowedCmds.Length > 0)
                            {
                                fallback.Commands.RemoveAll(c => !IsCommandAllowedByList(c, allowedCmds));
                            }
                            else if (inferredValuesOnly)
                            {
                                fallback.Commands.RemoveAll(c => c is not SetValuesCommand);
                            }
                            else if (inferredNoTitles)
                            {
                                fallback.Commands.RemoveAll(c => c is SetTitleCommand);
                            }
                            if (context.SelectionHardMode)
                            {
                                ApplySelectionHardMode(context, fallback);
                            }
                            return fallback;
                        }
                        catch { return new AIPlan(); }
                    }
                    string json2 = call2.Content ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(json2))
                    {
                        try
                        {
                            var doc2 = JsonDocument.Parse(json2);
                            var plan2 = ParsePlan(doc2);
                            plan2.Provider = string.IsNullOrWhiteSpace(call2.Provider) ? provider : call2.Provider;
                            plan2.Model = call2.Model;
                            plan2.Usage = call2.Usage;
                            plan2.LatencyMs = call2.LatencyMs;
                            if (allowedCmds != null && allowedCmds.Length > 0)
                            {
                                plan2.Commands.RemoveAll(c => !IsCommandAllowedByList(c, allowedCmds));
                            }
                            else if (inferredValuesOnly)
                            {
                                plan2.Commands.RemoveAll(c => c is not SetValuesCommand);
                            }
                            else if (inferredNoTitles)
                            {
                                plan2.Commands.RemoveAll(c => c is SetTitleCommand);
                            }
                            plan2.RawJson = json2; plan2.RawUser = rev; plan2.RawSystem = sys;
                            return plan2;
                        }
                        catch { }
                    }
                }
                // Selection hard mode: drop any write commands that touch outside the selection
                if (context.SelectionHardMode)
                {
                    ApplySelectionHardMode(context, plan);
                }
                return plan;
            }
            catch (Exception exParse)
            {
                try { Console.Error.WriteLine($"[ProviderChatPlanner] Post-API parse/validation failed: {exParse.GetType().Name}: {exParse.Message}"); } catch { }
                // Try to extract JSON substring if model wrapped it
                int i1 = json.IndexOf('{');
                int i2 = json.LastIndexOf('}');
                if (i1 >= 0 && i2 > i1)
                {
                    try {
                        var s = json.Substring(i1, i2 - i1 + 1);
                        var doc = JsonDocument.Parse(s);
                        var plan = ParsePlan(doc);
                        plan.RawJson = s; plan.RawUser = usr; plan.RawSystem = sys;
                        // Preserve provider metadata from the successful API call
                        plan.Provider = string.IsNullOrWhiteSpace(call.Provider) ? provider : call.Provider;
                        plan.Model = call.Model;
                        plan.Usage = call.Usage;
                        plan.LatencyMs = call.LatencyMs;
                        return plan;
                    }
                    catch { }
                }
                // If parsing fails, return empty plan
                return new AIPlan();
            }
        }

        private static string? TryLoadSystemPromptFromFile()
        {
            try
            {
                string? path = Environment.GetEnvironmentVariable("SYSTEM_PROMPT_PATH");
                if (string.IsNullOrWhiteSpace(path)) path = Path.Combine("prompts", "system_planner_v2.md");
                if (!Path.IsPathRooted(path))
                {
                    try
                    {
                        string cwd = Environment.CurrentDirectory;
                        string candidate = Path.Combine(cwd, path);
                        if (File.Exists(candidate)) path = candidate;
                        else
                        {
                            string baseDir = AppContext.BaseDirectory ?? cwd;
                            candidate = Path.Combine(baseDir, path);
                            if (File.Exists(candidate)) path = candidate;
                        }
                    }
                    catch { }
                }
                if (!File.Exists(path)) return null;
                string text = File.ReadAllText(path);
                ActivePromptPath = path;
                ActivePromptVersion = null;
                try
                {
                    string name = Path.GetFileName(path);
                    var m = System.Text.RegularExpressions.Regex.Match(name, @"system_planner_v(\d+)\.md", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                    if (m.Success && int.TryParse(m.Groups[1].Value, out int v)) ActivePromptVersion = v;
                }
                catch { }
                return string.IsNullOrWhiteSpace(text) ? null : text;
            }
            catch { return null; }
        }

        private static void ApplySelectionHardMode(AIContext ctx, AIPlan plan)
        {
            int r0 = Math.Max(0, ctx.StartRow);
            int c0 = Math.Max(0, ctx.StartCol);
            int r1 = r0 + Math.Max(1, ctx.Rows) - 1;
            int c1 = c0 + Math.Max(1, ctx.Cols) - 1;
            bool IsOob(int row, int col) => row < r0 || row > r1 || col < c0 || col > c1;
            plan.Commands.RemoveAll(cmd =>
            {
                if (cmd is SetValuesCommand sv)
                {
                    int rows = sv.Values.Length; int cols = rows > 0 ? (sv.Values[0]?.Length ?? 0) : 0;
                    for (int r = 0; r < rows; r++)
                    {
                        for (int c = 0; c < (sv.Values[r]?.Length ?? 0); c++)
                        {
                            if (IsOob(sv.StartRow + r, sv.StartCol + c)) return true;
                        }
                    }
                    return false;
                }
                else if (cmd is SetFormulaCommand sf)
                {
                    int rows = sf.Formulas.Length; int cols = rows > 0 ? (sf.Formulas[0]?.Length ?? 0) : 0;
                    int a1 = sf.StartRow; int b1 = sf.StartCol; int a2 = a1 + rows - 1; int b2 = b1 + cols - 1;
                    return (a1 < r0 || a2 > r1 || b1 < c0 || b2 > c1);
                }
                return false; // other command types allowed
            });
        }

        private static string[]? InferAllowedCommandsFromPrompt(string prompt)
        {
            // Lowercase for simple keyword checks
            var p = (prompt ?? string.Empty).ToLowerInvariant();
            var list = new System.Collections.Generic.List<string>();

            // Columns
            bool mentionsInsertCol = p.Contains("insert column") || p.Contains("add column") || p.Contains("insert col ") || p.Contains("insert col:") || p.Contains("insert into column");
            bool mentionsDeleteCol = p.Contains("delete column") || p.Contains("remove column") || p.Contains("del column") || p.Contains("del col");
            if (mentionsInsertCol) list.Add("insert_cols");
            if (mentionsDeleteCol) list.Add("delete_cols");

            // Rows
            bool mentionsInsertRow = p.Contains("insert row") || p.Contains("add row") || p.Contains("insert rows");
            bool mentionsDeleteRow = p.Contains("delete row") || p.Contains("remove row") || p.Contains("delete rows");
            if (mentionsInsertRow) list.Add("insert_rows");
            if (mentionsDeleteRow) list.Add("delete_rows");

            // Sort
            if (p.Contains("sort ") || p.Contains("alphabetize") || p.Contains("order by")) list.Add("sort_range");

            // Sheet deletion
            if (p.Contains("delete sheet") || p.Contains("remove sheet")) list.Add("delete_sheet");

            // Copy / Move
            if (p.Contains("copy ") && p.Contains(" to ")) list.Add("copy_range");
            if (p.Contains("move ") && p.Contains(" to ")) list.Add("move_range");

            // Formatting tasks
            if (p.Contains("bold ") || p.Contains("format ") || p.Contains("align ") || p.Contains("color ") || p.Contains("number format")) list.Add("set_format");

            // Validation tasks
            if (p.Contains("dropdown") || p.Contains("data validation") || p.Contains("allowed values") || p.Contains("between ")) list.Add("set_validation");

            // Conditional formatting
            if (p.Contains("highlight ") || p.Contains("conditional format") || p.Contains(">=") || p.Contains("<=") || p.Contains("threshold")) list.Add("set_conditional_format");

            // Deterministic transforms
            bool mentionsTrim = p.Contains("trim ") || p.Contains("remove extra spaces") || p.Contains("strip whitespace");
            bool mentionsUpper = p.Contains("uppercase") || p.Contains("upper case") || p.Contains("to upper");
            bool mentionsLower = p.Contains("lowercase") || p.Contains("lower case") || p.Contains("to lower");
            bool mentionsProper = p.Contains("proper case") || p.Contains("title case") || p.Contains("capitalize");
            bool mentionsPunct = p.Contains("strip punctuation") || p.Contains("remove punctuation");
            bool mentionsCity = p.Contains("normalize city") || p.Contains("clean city") || p.Contains("city names");
            if (mentionsTrim || mentionsUpper || mentionsLower || mentionsProper || mentionsPunct || mentionsCity) list.Add("transform_range");

            // If we inferred nothing structural, return null to keep default behavior
            return list.Count > 0 ? list.ToArray() : null;
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

            // Row-to-input mapping: explicitly map each input to its output row
            if (nearby != null && nearby.Length > 1)
            {
                int startRow1 = context.StartRow + 1; // 1-based starting row for the selection
                int fillColCount = Math.Max(0, headerRow.Length - 1);
                string fillStart = CellAddress.ColumnIndexToName(tableLeftCol + 1);
                string fillEnd = CellAddress.ColumnIndexToName(tableLeftCol + fillColCount);
                sb.Append($" IMPORTANT: Produce exactly one row of {fillColCount} values per input below, in the SAME order.");
                sb.Append($" Each output row MUST correspond to the input on that row. Do NOT reorder or skip inputs.");
                int inputCount = 0;
                for (int r = 1; r < nearby.Length; r++)
                {
                    if (nearby[r].Length > 0 && !string.IsNullOrWhiteSpace(nearby[r][0]))
                    {
                        inputCount++;
                        sb.Append($" Row {startRow1 + r - 1}: input=\"{nearby[r][0]}\" -> produce values for columns {fillStart}-{fillEnd} matching the headers.");
                    }
                }
                if (inputCount > 0)
                {
                    sb.Append($" Total expected rows in set_values: {inputCount}. Total columns per row: {fillColCount}.");
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
                        string typeDisplay = col.Type;
                        try
                        {
                            if (!string.IsNullOrWhiteSpace(col.Name) && col.Name.IndexOf("transliteration", System.StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                typeDisplay = string.IsNullOrWhiteSpace(typeDisplay) ? "text, latin alphabet" : (typeDisplay + ", latin alphabet");
                            }
                        }
                        catch { }
                        sb.Append(" (").Append(typeDisplay)
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
                InsertRowsCommand => "insert_rows",
                DeleteRowsCommand => "delete_rows",
                InsertColsCommand => "insert_cols",
                DeleteColsCommand => "delete_cols",
                DeleteSheetCommand => "delete_sheet",
                CopyRangeCommand => "copy_range",
                MoveRangeCommand => "move_range",
                SetFormatCommand => "set_format",
                SetValidationCommand => "set_validation",
                SetConditionalFormatCommand => "set_conditional_format",
                TransformRangeCommand => "transform_range",
                _ => string.Empty
            };
            if (string.IsNullOrEmpty(type)) return false;
            foreach (var a in allowed)
            {
                if (string.Equals(a?.Trim(), type, StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }

        private sealed class ChatCallResult
        {
            public string Content { get; set; } = string.Empty;
            public string Provider { get; set; } = string.Empty;
            public string? Model { get; set; }
            public AIUsage? Usage { get; set; }
            public int LatencyMs { get; set; }
        }

        private static async System.Threading.Tasks.Task<ChatCallResult> CallOpenAIWithUsageAsync(string system, string user, System.Threading.CancellationToken ct)
        {
            var key = Environment.GetEnvironmentVariable("OPENAI_API_KEY");
            if (string.IsNullOrWhiteSpace(key)) return new ChatCallResult { Provider = "OpenAI", Content = string.Empty };
            var model = Environment.GetEnvironmentVariable("OPENAI_MODEL") ?? "gpt-4o-mini";
            var endpoint = Environment.GetEnvironmentVariable("OPENAI_BASE_URL") ?? "https://api.openai.com/v1/chat/completions";
            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(20) };
            http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", key);
            http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            var body = new { model, temperature = 0.0, messages = new object[] { new { role = "system", content = system }, new { role = "user", content = user } } };
            using var req = new HttpRequestMessage(HttpMethod.Post, endpoint) { Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json") };
            var t0 = DateTime.UtcNow;
            using var resp = await http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false);
            resp.EnsureSuccessStatusCode();
            await using var stream = await resp.Content.ReadAsStreamAsync(ct).ConfigureAwait(false);
            using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: ct).ConfigureAwait(false);
            var result = new ChatCallResult { Provider = "OpenAI", Model = model, LatencyMs = (int)(DateTime.UtcNow - t0).TotalMilliseconds };
            try { if (doc.RootElement.TryGetProperty("model", out var mdl) && mdl.ValueKind == JsonValueKind.String) result.Model = mdl.GetString(); } catch { }
            try
            {
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
            if (doc.RootElement.TryGetProperty("choices", out var choices) && choices.ValueKind == JsonValueKind.Array && choices.GetArrayLength() > 0)
            {
                var msg = choices[0].GetProperty("message");
                if (msg.TryGetProperty("content", out var cnt))
                {
                    if (cnt.ValueKind == JsonValueKind.String) { result.Content = cnt.GetString() ?? string.Empty; return result; }
                    if (cnt.ValueKind == JsonValueKind.Array)
                    {
                        var sb = new StringBuilder();
                        foreach (var part in cnt.EnumerateArray()) if (part.TryGetProperty("text", out var t) && t.ValueKind == JsonValueKind.String) sb.Append(t.GetString());
                        result.Content = sb.ToString();
                        return result;
                    }
                }
            }
            result.Content = string.Empty; return result;
        }

        private static async System.Threading.Tasks.Task<ChatCallResult> CallAnthropicWithUsageAsync(string system, string user, System.Threading.CancellationToken ct)
        {
            var key = Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY");
            if (string.IsNullOrWhiteSpace(key)) return new ChatCallResult { Provider = "Anthropic", Content = string.Empty };
            var model = Environment.GetEnvironmentVariable("ANTHROPIC_MODEL") ?? "claude-3-haiku-20240307";
            var endpoint = Environment.GetEnvironmentVariable("ANTHROPIC_BASE_URL") ?? "https://api.anthropic.com/v1/messages";
            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(20) };
            http.DefaultRequestHeaders.Add("x-api-key", key);
            http.DefaultRequestHeaders.Add("anthropic-version", "2023-06-01");
            int maxTokens = 2048; try { var s = Environment.GetEnvironmentVariable("ANTHROPIC_MAX_TOKENS"); if (!string.IsNullOrWhiteSpace(s)) maxTokens = int.Parse(s); } catch { }
            var body = new { model, max_tokens = maxTokens, temperature = 0.0, system = system, messages = new object[] { new { role = "user", content = new object[] { new { type = "text", text = user } } } } };
            using var req = new HttpRequestMessage(HttpMethod.Post, endpoint) { Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json") };
            var t0 = DateTime.UtcNow;
            using var resp = await http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false);
            resp.EnsureSuccessStatusCode();
            await using var stream = await resp.Content.ReadAsStreamAsync(ct).ConfigureAwait(false);
            using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: ct).ConfigureAwait(false);
            var result = new ChatCallResult { Provider = "Anthropic", Model = model, LatencyMs = (int)(DateTime.UtcNow - t0).TotalMilliseconds };
            try { if (doc.RootElement.TryGetProperty("model", out var mdl) && mdl.ValueKind == JsonValueKind.String) result.Model = mdl.GetString(); } catch { }
            try
            {
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
                foreach (var part in contentEl.EnumerateArray()) { if (part.TryGetProperty("text", out var t) && t.ValueKind == JsonValueKind.String) { text = t.GetString(); break; } }
                if (!string.IsNullOrEmpty(text)) { result.Content = text!; return result; }
            }
            result.Content = string.Empty; return result;
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

        private static System.Collections.Generic.List<string> ValidatePlanAgainstContext(AIContext ctx, AIPlan plan, string[]? allowedCmds, bool inferredValuesOnly, bool inferredNoTitles, out int expectedRowWidth)
        {
            var issues = new System.Collections.Generic.List<string>();
            int r0 = Math.Max(0, ctx.StartRow);
            int c0 = Math.Max(0, ctx.StartCol);
            int r1 = r0 + Math.Max(1, ctx.Rows) - 1;
            int c1 = c0 + Math.Max(1, ctx.Cols) - 1;
            expectedRowWidth = Math.Max(1, ctx.Cols);

            // Build quick lookups
            var writable = new System.Collections.Generic.HashSet<int>();
            bool hasWritable = false;
            int? inputCol = null; bool allowInputExisting = false; bool allowInputEmpty = false;
            if (ctx.WritePolicy != null)
            {
                if (ctx.WritePolicy.WritableColumns != null && ctx.WritePolicy.WritableColumns.Length > 0)
                {
                    foreach (var w in ctx.WritePolicy.WritableColumns) { if (w >= 0) { writable.Add(w); hasWritable = true; } }
                }
                inputCol = ctx.WritePolicy.InputColumnIndex;
                allowInputExisting = ctx.WritePolicy.AllowInputWritesForExistingRows;
                allowInputEmpty = ctx.WritePolicy.AllowInputWritesForEmptyRows;
            }

            // Helper to check if a given (row) inside selection has an existing input value
            bool RowHasExistingInputInSelection(int rowZero)
            {
                try
                {
                    if (!inputCol.HasValue) return false;
                    if (ctx.SelectionValues == null) return false;
                    int selRow = rowZero - r0;
                    if (selRow < 0 || selRow >= ctx.SelectionValues.Length) return false;
                    int selCol = inputCol.Value - c0;
                    if (selCol < 0 || ctx.SelectionValues[selRow] == null) return false;
                    var line = ctx.SelectionValues[selRow];
                    if (selCol >= line.Length) return false;
                    var v = line[selCol] ?? string.Empty;
                    return !string.IsNullOrWhiteSpace(v);
                }
                catch { return false; }
            }

            // Validate each command
            foreach (var cmd in plan.Commands)
            {
                // Allowed commands guard (redundant with filter but adds explanation)
                if (allowedCmds != null && allowedCmds.Length > 0 && !IsCommandAllowedByList(cmd, allowedCmds))
                {
                    issues.Add($"Command not allowed by policy: {DescribeType(cmd)}");
                    continue;
                }
                if (inferredValuesOnly && cmd is not SetValuesCommand)
                {
                    issues.Add($"Only set_values allowed, found: {DescribeType(cmd)}");
                    continue;
                }
                if (inferredNoTitles && cmd is SetTitleCommand)
                {
                    issues.Add("set_title not allowed for this request");
                    continue;
                }

                if (cmd is SetValuesCommand sv)
                {
                    // Width check: require consistent width == expectedRowWidth
                    for (int r = 0; r < sv.Values.Length; r++)
                    {
                        int w = sv.Values[r]?.Length ?? 0;
                        if (w > 0 && w != expectedRowWidth)
                        {
                            issues.Add($"Row width {w} != expected {expectedRowWidth} in set_values");
                            break;
                        }
                    }
                    // Bounds + policy check (first violation only per category)
                    bool oobNoted = false; bool nonWritableNoted = false; bool inputPolicyNoted = false;
                    for (int r = 0; r < sv.Values.Length; r++)
                    {
                        int row = sv.StartRow + r;
                        int cols = sv.Values[r]?.Length ?? 0;
                        for (int c = 0; c < cols; c++)
                        {
                            int col = sv.StartCol + c;
                            if (!oobNoted && (row < r0 || row > r1 || col < c0 || col > c1))
                            {
                                issues.Add($"Writes outside selection at {SpreadsheetApp.Core.CellAddress.ToAddress(row, col)}");
                                oobNoted = true;
                            }
                            if (!nonWritableNoted && hasWritable && !writable.Contains(col))
                            {
                                issues.Add($"Writes to non-writable column {SpreadsheetApp.Core.CellAddress.ColumnIndexToName(col)}");
                                nonWritableNoted = true;
                            }
                            if (!inputPolicyNoted && inputCol.HasValue && col == inputCol.Value)
                            {
                                bool hasExisting = RowHasExistingInputInSelection(row);
                                if ((hasExisting && !allowInputExisting) || (!hasExisting && !allowInputEmpty))
                                {
                                    issues.Add($"Writes to input column {SpreadsheetApp.Core.CellAddress.ColumnIndexToName(col)} in a row where policy forbids it");
                                    inputPolicyNoted = true;
                                }
                            }
                            if (oobNoted && nonWritableNoted && inputPolicyNoted) break;
                        }
                        if (oobNoted && nonWritableNoted && inputPolicyNoted) break;
                    }
                }
                else if (cmd is SetFormulaCommand sf)
                {
                    int rows = sf.Formulas.Length; int cols = rows > 0 ? sf.Formulas[0].Length : 0;
                    int a1 = sf.StartRow; int b1 = sf.StartCol; int a2 = a1 + rows - 1; int b2 = b1 + cols - 1;
                    if (a1 < r0 || a2 > r1 || b1 < c0 || b2 > c1)
                    {
                        issues.Add("set_formula writes outside selection");
                    }
                    if (hasWritable)
                    {
                        for (int c = 0; c < cols; c++)
                        {
                            int abs = b1 + c; if (!writable.Contains(abs)) { issues.Add($"set_formula uses non-writable column {SpreadsheetApp.Core.CellAddress.ColumnIndexToName(abs)}"); break; }
                        }
                    }
                }
                else if (cmd is ClearRangeCommand cr)
                {
                    int a1 = cr.StartRow; int b1 = cr.StartCol; int a2 = a1 + Math.Max(1, cr.Rows) - 1; int b2 = b1 + Math.Max(1, cr.Cols) - 1;
                    if (a1 < r0 || a2 > r1 || b1 < c0 || b2 > c1)
                    {
                        issues.Add("clear_range outside selection");
                    }
                }
                else if (cmd is SortRangeCommand sr)
                {
                    int a1 = sr.StartRow; int b1 = sr.StartCol; int a2 = a1 + Math.Max(1, sr.Rows) - 1; int b2 = b1 + Math.Max(1, sr.Cols) - 1;
                    if (a1 < r0 || a2 > r1 || b1 < c0 || b2 > c1)
                    {
                        issues.Add("sort_range outside selection");
                    }
                }
            }
            // De-duplicate messages to keep the revision prompt short
            var dedup = new System.Collections.Generic.HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
            var trimmed = new System.Collections.Generic.List<string>();
            foreach (var i in issues) if (dedup.Add(i)) trimmed.Add(i);
            // Limit to a handful to avoid long prompts
            if (trimmed.Count > 6) trimmed = trimmed.GetRange(0, 6);
            return trimmed;
        }

        private static string DescribeType(IAICommand cmd)
        {
            return cmd switch
            {
                SetValuesCommand => "set_values",
                SetFormulaCommand => "set_formula",
                SetTitleCommand => "set_title",
                CreateSheetCommand => "create_sheet",
                ClearRangeCommand => "clear_range",
                RenameSheetCommand => "rename_sheet",
                SortRangeCommand => "sort_range",
                InsertRowsCommand => "insert_rows",
                DeleteRowsCommand => "delete_rows",
                _ => cmd.GetType().Name
            };
        }

        private static string BuildRevisionUserMessage(AIContext ctx, string originalUser, string[]? allowedCmds, int expectedWidth, System.Collections.Generic.List<string> violations, string priorPlanJson)
        {
            var sb = new StringBuilder();
            // Carry the original context/user string for grounding
            sb.Append(originalUser);
            sb.Append(' ');
            sb.Append("REVISION REQUEST: The previous plan violated constraints. Fix the issues below and return a corrected plan as strict JSON only. ");
            // Constraints summary
            sb.Append($"Selection bounds: start=({ctx.StartRow + 1},{SpreadsheetApp.Core.CellAddress.ColumnIndexToName(ctx.StartCol)}), Rows={Math.Max(1, ctx.Rows)}, Cols={Math.Max(1, ctx.Cols)}. ");
            if (allowedCmds != null && allowedCmds.Length > 0)
            {
                sb.Append("Allowed command types: ");
                for (int i = 0; i < allowedCmds.Length; i++) { if (i > 0) sb.Append(','); sb.Append(allowedCmds[i]); }
                sb.Append(". ");
            }
            // Writable columns summary
            if (ctx.WritePolicy?.WritableColumns != null && ctx.WritePolicy.WritableColumns.Length > 0)
            {
                sb.Append("Writable columns: ");
                for (int i = 0; i < ctx.WritePolicy.WritableColumns.Length; i++)
                {
                    if (i > 0) sb.Append(',');
                    sb.Append(SpreadsheetApp.Core.CellAddress.ColumnIndexToName(ctx.WritePolicy.WritableColumns[i]));
                }
                sb.Append(". ");
            }
            if (ctx.WritePolicy?.InputColumnIndex != null)
            {
                string letter = SpreadsheetApp.Core.CellAddress.ColumnIndexToName(ctx.WritePolicy.InputColumnIndex.Value);
                if (!ctx.WritePolicy.AllowInputWritesForExistingRows && !ctx.WritePolicy.AllowInputWritesForEmptyRows)
                    sb.Append($"Do not write to input column {letter}. ");
                else if (!ctx.WritePolicy.AllowInputWritesForExistingRows && ctx.WritePolicy.AllowInputWritesForEmptyRows)
                    sb.Append($"Do not modify existing values in input column {letter}; writing new inputs only to empty rows in selection is allowed. ");
            }
            // Emphasize append-only and selection-row-only behaviors when applicable
            try
            {
                int r1 = ctx.StartRow + 1;
                int r2 = ctx.StartRow + Math.Max(1, ctx.Rows);
                if (r2 >= r1)
                {
                    sb.Append($"Write outputs only for rows {r1}-{r2} within the selection; do not modify other rows. ");
                }
                if (ctx.WritePolicy?.InputColumnIndex != null && ctx.WritePolicy.AllowInputWritesForEmptyRows && !ctx.WritePolicy.AllowInputWritesForExistingRows)
                {
                    string letter = SpreadsheetApp.Core.CellAddress.ColumnIndexToName(ctx.WritePolicy.InputColumnIndex.Value);
                    sb.Append($"Append-only for input column {letter}: add new inputs only in empty cells within rows {r1}-{r2}; do not change existing inputs.");
                    sb.Append(' ');
                }
            }
            catch { }
            sb.Append($"Expected per-row width: {expectedWidth}. ");
            // Violations list
            sb.Append("Problems: ");
            for (int i = 0; i < violations.Count; i++) { if (i > 0) sb.Append(" | "); sb.Append(violations[i]); }
            sb.Append(". Reissue the full corrected plan now. No prose, no extra keys, JSON only.");
            // Include previous plan JSON at the end for reference (helps some models repair precisely)
            if (!string.IsNullOrWhiteSpace(priorPlanJson))
            {
                sb.Append(" PreviousPlan= ");
                sb.Append(priorPlanJson);
            }
            return sb.ToString();
        }

        private static AIPlan ParsePlan(JsonDocument doc)
        {
            var plan = new AIPlan();
            if (doc.RootElement.ValueKind != JsonValueKind.Object) return plan;
            // Queries (optional)
            if (doc.RootElement.TryGetProperty("queries", out var queries) && queries.ValueKind == JsonValueKind.Array)
            {
                foreach (var q in queries.EnumerateArray())
                {
                    try
                    {
                        if (!q.TryGetProperty("type", out var qt)) continue;
                        string? qtype = qt.GetString()?.Trim().ToLowerInvariant();
                        switch (qtype)
                        {
                            case "selection_summary":
                                plan.Queries.Add(new SelectionSummaryQuery());
                                break;
                            case "profile_column":
                                {
                                    var pq = new ProfileColumnQuery();
                                    if (q.TryGetProperty("col", out var colEl)) pq.ColumnIndex = ParseColumnIndex(colEl);
                                    if (q.TryGetProperty("rows", out var rr) && rr.ValueKind == JsonValueKind.Number) pq.Rows = SafeInt(rr, 0);
                                    plan.Queries.Add(pq);
                                    break;
                                }
                            case "describe_column":
                                {
                                    var dq = new DescribeColumnQuery();
                                    if (q.TryGetProperty("col", out var colEl)) dq.ColumnIndex = ParseColumnIndex(colEl);
                                    if (q.TryGetProperty("rows", out var rr) && rr.ValueKind == JsonValueKind.Number) dq.Rows = SafeInt(rr, 0);
                                    plan.Queries.Add(dq);
                                    break;
                                }
                            case "unique_values":
                                {
                                    var uq = new UniqueValuesQuery();
                                    if (q.TryGetProperty("col", out var colEl)) uq.ColumnIndex = ParseColumnIndex(colEl);
                                    if (q.TryGetProperty("top", out var top) && top.ValueKind == JsonValueKind.Number) uq.TopK = SafeInt(top, 10);
                                    plan.Queries.Add(uq);
                                    break;
                                }
                            case "sample_rows":
                                {
                                    var sq = new SampleRowsQuery();
                                    if (q.TryGetProperty("rows", out var r) && r.ValueKind == JsonValueKind.Number) sq.Rows = SafeInt(r, 5);
                                    if (q.TryGetProperty("cols", out var c) && c.ValueKind == JsonValueKind.Number) sq.Cols = SafeInt(c, 6);
                                    plan.Queries.Add(sq);
                                    break;
                                }
                            case "count_where":
                                {
                                    var cq = new CountWhereQuery();
                                    if (q.TryGetProperty("filters", out var fs) && fs.ValueKind == JsonValueKind.Array)
                                    {
                                        foreach (var f in fs.EnumerateArray())
                                        {
                                            try
                                            {
                                                var filter = new CountWhereQuery.Filter();
                                                if (f.TryGetProperty("col", out var colE)) filter.ColumnIndex = ParseColumnIndex(colE);
                                                if (f.TryGetProperty("op", out var opE) && opE.ValueKind == JsonValueKind.String) filter.Op = opE.GetString() ?? "eq";
                                                if (f.TryGetProperty("value", out var valE)) filter.Value = valE.ValueKind == JsonValueKind.String ? (valE.GetString() ?? string.Empty) : valE.GetRawText();
                                                cq.Filters.Add(filter);
                                            }
                                            catch { }
                                        }
                                    }
                                    plan.Queries.Add(cq);
                                    break;
                                }
                        }
                    }
                    catch { }
                }
            }
            // Commands (optional)
            if (!doc.RootElement.TryGetProperty("commands", out var cmds) || cmds.ValueKind != JsonValueKind.Array) return plan;
            foreach (var cmd in cmds.EnumerateArray())
            {
                try
                {
                    if (!cmd.TryGetProperty("type", out var t)) continue;
                    var type = t.GetString()?.Trim().ToLowerInvariant();
                    switch (type)
                    {
                        case "transform_range":
                            {
                                var tr = new TransformRangeCommand();
                                if (cmd.TryGetProperty("start", out var start)) { int srr, scc; ParseStart(start, out srr, out scc); tr.StartRow = srr; tr.StartCol = scc; }
                                if (cmd.TryGetProperty("rows", out var r)) tr.Rows = SafeInt(r, 1);
                                if (cmd.TryGetProperty("cols", out var c)) tr.Cols = SafeInt(c, 1);
                                if (cmd.TryGetProperty("op", out var op) && op.ValueKind == JsonValueKind.String) tr.Op = op.GetString() ?? "trim";
                                if (cmd.TryGetProperty("rationale", out var rat) && rat.ValueKind == JsonValueKind.String) tr.Rationale = rat.GetString();
                                plan.Commands.Add(tr);
                                break;
                            }
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
                            if (cmd.TryGetProperty("rationale", out var rsv) && rsv.ValueKind == JsonValueKind.String) sv.Rationale = rsv.GetString();
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
                            if (cmd.TryGetProperty("rationale", out var rsf) && rsf.ValueKind == JsonValueKind.String) sf.Rationale = rsf.GetString();
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
                                if (cmd.TryGetProperty("rationale", out var rcr) && rcr.ValueKind == JsonValueKind.String) cr.Rationale = rcr.GetString();
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
                                if (cmd.TryGetProperty("rationale", out var rst) && rst.ValueKind == JsonValueKind.String) st.Rationale = rst.GetString();
                                plan.Commands.Add(st);
                                break;
                            }
                        case "rename_sheet":
                            {
                                var rn = new RenameSheetCommand();
                                if (cmd.TryGetProperty("index", out var idx)) rn.Index1 = SafeInt(idx, 0) > 0 ? SafeInt(idx, 0) : (int?)null;
                                if (cmd.TryGetProperty("old_name", out var on) && on.ValueKind == JsonValueKind.String) rn.OldName = on.GetString();
                                if (cmd.TryGetProperty("new_name", out var nn) && nn.ValueKind == JsonValueKind.String) rn.NewName = nn.GetString() ?? rn.NewName;
                                if (cmd.TryGetProperty("rationale", out var rrn) && rrn.ValueKind == JsonValueKind.String) rn.Rationale = rrn.GetString();
                                plan.Commands.Add(rn);
                                break;
                            }
                        case "create_sheet":
                            {
                                var cs = new CreateSheetCommand();
                                if (cmd.TryGetProperty("name", out var n)) cs.Name = n.GetString() ?? cs.Name;
                                if (cmd.TryGetProperty("rationale", out var rcs) && rcs.ValueKind == JsonValueKind.String) cs.Rationale = rcs.GetString();
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
                                if (cmd.TryGetProperty("rationale", out var rsr) && rsr.ValueKind == JsonValueKind.String) sr.Rationale = rsr.GetString();
                                plan.Commands.Add(sr);
                                break;
                            }
                        case "insert_rows":
                            {
                                var ir = new InsertRowsCommand();
                                if (cmd.TryGetProperty("at", out var at)) ir.At = Math.Max(0, SafeInt(at, 1) - 1);
                                if (cmd.TryGetProperty("count", out var cnt)) ir.Count = Math.Max(1, SafeInt(cnt, 1));
                                if (cmd.TryGetProperty("rationale", out var rir) && rir.ValueKind == JsonValueKind.String) ir.Rationale = rir.GetString();
                                plan.Commands.Add(ir);
                                break;
                            }
                        case "delete_rows":
                            {
                                var dr = new DeleteRowsCommand();
                                if (cmd.TryGetProperty("at", out var at)) dr.At = Math.Max(0, SafeInt(at, 1) - 1);
                                if (cmd.TryGetProperty("count", out var cnt)) dr.Count = Math.Max(1, SafeInt(cnt, 1));
                                if (cmd.TryGetProperty("rationale", out var rdr) && rdr.ValueKind == JsonValueKind.String) dr.Rationale = rdr.GetString();
                                plan.Commands.Add(dr);
                                break;
                            }
                        case "set_validation":
                            {
                                var sv = new SetValidationCommand();
                                if (cmd.TryGetProperty("start", out var start)) { int r0, c0; ParseStart(start, out r0, out c0); sv.StartRow = r0; sv.StartCol = c0; }
                                if (cmd.TryGetProperty("rows", out var rr)) sv.Rows = SafeInt(rr, 1);
                                if (cmd.TryGetProperty("cols", out var cc)) sv.Cols = SafeInt(cc, 1);
                                if (cmd.TryGetProperty("mode", out var md) && md.ValueKind == JsonValueKind.String) sv.Mode = md.GetString() ?? "none";
                                if (cmd.TryGetProperty("allow_empty", out var ae) && (ae.ValueKind == JsonValueKind.True || ae.ValueKind == JsonValueKind.False)) sv.AllowEmpty = ae.GetBoolean();
                                if (cmd.TryGetProperty("min", out var mn)) { try { sv.Min = mn.GetDouble(); } catch { } }
                                if (cmd.TryGetProperty("max", out var mx)) { try { sv.Max = mx.GetDouble(); } catch { } }
                                if (cmd.TryGetProperty("allowed", out var al) && al.ValueKind == JsonValueKind.Array)
                                {
                                    var list = new System.Collections.Generic.List<string>();
                                    foreach (var el in al.EnumerateArray()) if (el.ValueKind == JsonValueKind.String) list.Add(el.GetString() ?? string.Empty);
                                    sv.AllowedList = list.ToArray();
                                }
                                plan.Commands.Add(sv);
                                break;
                            }
                        case "set_conditional_format":
                            {
                                var scf = new SetConditionalFormatCommand();
                                if (cmd.TryGetProperty("start", out var start)) { int r0, c0; ParseStart(start, out r0, out c0); scf.StartRow = r0; scf.StartCol = c0; }
                                if (cmd.TryGetProperty("rows", out var rr)) scf.Rows = SafeInt(rr, 1);
                                if (cmd.TryGetProperty("cols", out var cc)) scf.Cols = SafeInt(cc, 1);
                                if (cmd.TryGetProperty("op", out var op) && op.ValueKind == JsonValueKind.String) scf.Operator = op.GetString() ?? ">";
                                if (cmd.TryGetProperty("threshold", out var th)) { try { scf.Threshold = th.GetDouble(); } catch { } }
                                if (cmd.TryGetProperty("bold", out var b) && (b.ValueKind == JsonValueKind.True || b.ValueKind == JsonValueKind.False)) scf.Bold = b.GetBoolean();
                                if (cmd.TryGetProperty("number_format", out var nf) && nf.ValueKind == JsonValueKind.String) scf.NumberFormat = nf.GetString();
                                if (cmd.TryGetProperty("h_align", out var ha) && ha.ValueKind == JsonValueKind.String) scf.HAlign = ha.GetString();
                                if (cmd.TryGetProperty("fore_color", out var fc) && fc.ValueKind == JsonValueKind.String) scf.ForeColorArgb = ParseHexColor(fc.GetString());
                                if (cmd.TryGetProperty("back_color", out var bc) && bc.ValueKind == JsonValueKind.String) scf.BackColorArgb = ParseHexColor(bc.GetString());
                                plan.Commands.Add(scf);
                                break;
                            }
                        case "copy_range":
                            {
                                var cr = new CopyRangeCommand();
                                if (cmd.TryGetProperty("start", out var start)) { int r0, c0; ParseStart(start, out r0, out c0); cr.StartRow = r0; cr.StartCol = c0; }
                                if (cmd.TryGetProperty("rows", out var rr)) cr.Rows = SafeInt(rr, 1);
                                if (cmd.TryGetProperty("cols", out var cc)) cr.Cols = SafeInt(cc, 1);
                                if (cmd.TryGetProperty("dest", out var dest)) { int rd, cd; ParseStart(dest, out rd, out cd); cr.DestRow = rd; cr.DestCol = cd; }
                                if (cmd.TryGetProperty("rationale", out var rcp) && rcp.ValueKind == JsonValueKind.String) cr.Rationale = rcp.GetString();
                                plan.Commands.Add(cr);
                                break;
                            }
                        case "move_range":
                            {
                                var mr = new MoveRangeCommand();
                                if (cmd.TryGetProperty("start", out var start)) { int r0, c0; ParseStart(start, out r0, out c0); mr.StartRow = r0; mr.StartCol = c0; }
                                if (cmd.TryGetProperty("rows", out var rr)) mr.Rows = SafeInt(rr, 1);
                                if (cmd.TryGetProperty("cols", out var cc)) mr.Cols = SafeInt(cc, 1);
                                if (cmd.TryGetProperty("dest", out var dest)) { int rd, cd; ParseStart(dest, out rd, out cd); mr.DestRow = rd; mr.DestCol = cd; }
                                if (cmd.TryGetProperty("rationale", out var rmv) && rmv.ValueKind == JsonValueKind.String) mr.Rationale = rmv.GetString();
                                plan.Commands.Add(mr);
                                break;
                            }
                        case "set_format":
                            {
                                var sfmt = new SetFormatCommand();
                                if (cmd.TryGetProperty("start", out var start)) { int r0, c0; ParseStart(start, out r0, out c0); sfmt.StartRow = r0; sfmt.StartCol = c0; }
                                if (cmd.TryGetProperty("rows", out var rr)) sfmt.Rows = SafeInt(rr, 1);
                                if (cmd.TryGetProperty("cols", out var cc)) sfmt.Cols = SafeInt(cc, 1);
                                if (cmd.TryGetProperty("bold", out var b) && (b.ValueKind == JsonValueKind.True || b.ValueKind == JsonValueKind.False)) sfmt.Bold = b.GetBoolean();
                                if (cmd.TryGetProperty("number_format", out var nf) && nf.ValueKind == JsonValueKind.String) { var s = nf.GetString(); if (!string.IsNullOrWhiteSpace(s)) sfmt.NumberFormat = s; }
                                if (cmd.TryGetProperty("h_align", out var ha) && ha.ValueKind == JsonValueKind.String) { var s = ha.GetString(); if (!string.IsNullOrWhiteSpace(s)) sfmt.HAlign = s!.Trim().ToLowerInvariant(); }
                                if (cmd.TryGetProperty("fore_color", out var fc) && fc.ValueKind == JsonValueKind.String) sfmt.ForeColorArgb = ParseHexColor(fc.GetString());
                                if (cmd.TryGetProperty("back_color", out var bc) && bc.ValueKind == JsonValueKind.String) sfmt.BackColorArgb = ParseHexColor(bc.GetString());
                                if (cmd.TryGetProperty("rationale", out var rfmt) && rfmt.ValueKind == JsonValueKind.String) sfmt.Rationale = rfmt.GetString();
                                plan.Commands.Add(sfmt);
                                break;
                            }
                        case "insert_cols":
                            {
                                var ic = new InsertColsCommand();
                                if (cmd.TryGetProperty("at", out var at))
                                {
                                    if (at.ValueKind == JsonValueKind.String)
                                        ic.At = Math.Max(0, CellAddress.ColumnNameToIndex(at.GetString() ?? "A"));
                                    else
                                        ic.At = Math.Max(0, SafeInt(at, 1) - 1);
                                }
                                if (cmd.TryGetProperty("count", out var cnt)) ic.Count = Math.Max(1, SafeInt(cnt, 1));
                                if (cmd.TryGetProperty("rationale", out var ric) && ric.ValueKind == JsonValueKind.String) ic.Rationale = ric.GetString();
                                plan.Commands.Add(ic);
                                break;
                            }
                        case "delete_cols":
                            {
                                var dc = new DeleteColsCommand();
                                if (cmd.TryGetProperty("at", out var at))
                                {
                                    if (at.ValueKind == JsonValueKind.String)
                                        dc.At = Math.Max(0, CellAddress.ColumnNameToIndex(at.GetString() ?? "A"));
                                    else
                                        dc.At = Math.Max(0, SafeInt(at, 1) - 1);
                                }
                                if (cmd.TryGetProperty("count", out var cnt)) dc.Count = Math.Max(1, SafeInt(cnt, 1));
                                if (cmd.TryGetProperty("rationale", out var rdc) && rdc.ValueKind == JsonValueKind.String) dc.Rationale = rdc.GetString();
                                plan.Commands.Add(dc);
                                break;
                            }
                        case "delete_sheet":
                            {
                                var ds = new DeleteSheetCommand();
                                if (cmd.TryGetProperty("index", out var idx)) ds.Index1 = SafeInt(idx, 0) > 0 ? SafeInt(idx, 0) : (int?)null;
                                if (cmd.TryGetProperty("name", out var nm) && nm.ValueKind == JsonValueKind.String) ds.Name = nm.GetString();
                                if (cmd.TryGetProperty("rationale", out var rds) && rds.ValueKind == JsonValueKind.String) ds.Rationale = rds.GetString();
                                plan.Commands.Add(ds);
                                break;
                            }
                    }
                }
                catch { }
            }
            return plan;
        }

        private static int ParseColumnIndex(JsonElement el)
        {
            // Accept a string letter (e.g., "B") or 1-based numeric index
            try
            {
                if (el.ValueKind == JsonValueKind.String)
                {
                    string s = el.GetString() ?? string.Empty;
                    if (SpreadsheetApp.Core.CellAddress.TryParse((s + "1").ToUpperInvariant(), out int _, out int col)) return col;
                }
                else if (el.ValueKind == JsonValueKind.Number)
                {
                    int idx = SafeInt(el, 1) - 1;
                    return Math.Max(0, idx);
                }
            }
            catch { }
            return 0;
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

        private static int? ParseHexColor(string? s)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(s)) return null;
                s = s.Trim();
                if (s.StartsWith("#")) s = s.Substring(1);
                if (s.Length == 6)
                {
                    // RRGGBB -> ARGB with A=FF
                    int rgb = Convert.ToInt32(s, 16);
                    return unchecked((int)(0xFF000000u | (uint)rgb));
                }
                if (s.Length == 8)
                {
                    // AARRGGBB
                    int argb = Convert.ToInt32(s, 16);
                    return argb;
                }
                return null;
            }
            catch { return null; }
        }
    }
}
