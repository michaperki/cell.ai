using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace SpreadsheetApp.Core.AI
{
    public interface IChatPlanner
    {
        Task<AIPlan> PlanAsync(AIContext context, string prompt, CancellationToken cancellationToken);
    }

    public sealed class MockChatPlanner : IChatPlanner
    {
        public Task<AIPlan> PlanAsync(AIContext context, string prompt, CancellationToken cancellationToken)
        {
            prompt = (prompt ?? string.Empty).Trim().ToLowerInvariant();
            var plan = new AIPlan();
            plan.RawUser = $"Sheet={context.SheetName}; Selection=({context.StartRow+1},{SpreadsheetApp.Core.CellAddress.ColumnIndexToName(context.StartCol)}); Rows={context.Rows}; Cols={context.Cols}; Title={(context.Title??string.Empty)}. Instruction={(prompt)}.";
            plan.RawSystem = "Mock planner: deterministic heuristics without external provider.";

            // If prompt suggests a list, generate a set-values plan similar to Generate Fill
            bool wantsList = prompt.Contains("list") || prompt.Contains("grocery") || prompt.Contains("todo") || prompt.Contains("ingredients") || prompt.Contains("caesar");
            if (wantsList)
            {
                // Derive size hints
                int rows = Math.Max(5, context.Rows);
                int cols = Math.Max(1, context.Cols);
                // Parse NxM patterns
                var nxm = System.Text.RegularExpressions.Regex.Match(prompt, @"(\d+)\s*[xX]\s*(\d+)");
                if (nxm.Success)
                {
                    int.TryParse(nxm.Groups[1].Value, out rows);
                    int.TryParse(nxm.Groups[2].Value, out cols);
                }
                // Parse "N rows" / "N columns"
                var rmatch = System.Text.RegularExpressions.Regex.Match(prompt, @"(\d+)\s+row");
                if (rmatch.Success) int.TryParse(rmatch.Groups[1].Value, out rows);
                var cmatch = System.Text.RegularExpressions.Regex.Match(prompt, @"(\d+)\s+col");
                if (cmatch.Success) int.TryParse(cmatch.Groups[1].Value, out cols);
                // Parse words: two/three/four columns
                if (prompt.Contains("two column")) cols = Math.Max(cols, 2);
                if (prompt.Contains("three column")) cols = Math.Max(cols, 3);
                if (prompt.Contains("four column")) cols = Math.Max(cols, 4);

                // Parse target column letter
                int startCol = context.StartCol;
                var colMatch = System.Text.RegularExpressions.Regex.Match(prompt, @"column\s+([a-zA-Z]+)");
                if (colMatch.Success)
                {
                    try { startCol = SpreadsheetApp.Core.CellAddress.ColumnNameToIndex(colMatch.Groups[1].Value); } catch { }
                }
                // Parse target row number
                int startRow = context.StartRow;
                var rowMatch = System.Text.RegularExpressions.Regex.Match(prompt, @"row\s+(\d+)");
                if (rowMatch.Success) { if (int.TryParse(rowMatch.Groups[1].Value, out int r)) startRow = Math.Max(0, r - 1); }
                // Append semantics: continue/add/append/below -> mark with sentinel
                if (prompt.Contains("append") || prompt.Contains("continue") || prompt.Contains("add more") || prompt.Contains("below"))
                {
                    startRow = -1; // sentinel handled by ApplyPlan
                }

                // Reuse MockInferenceProvider heuristics
                var inf = new MockInferenceProvider();
                var infCtx = new AIContext
                {
                    SheetName = context.SheetName,
                    StartRow = startRow,
                    StartCol = startCol,
                    Rows = rows,
                    Cols = cols,
                    Title = context.Title,
                    Prompt = prompt
                };
                var res = inf.GenerateFillAsync(infCtx, CancellationToken.None).Result;
                plan.Commands.Add(new SetValuesCommand { StartRow = startRow, StartCol = startCol, Values = res.Cells });
                return Task.FromResult(plan);
            }

            // Create sheet if requested
            if (prompt.Contains("create sheet") || prompt.Contains("new sheet"))
            {
                string name = "Sheet" + DateTime.Now.ToString("HHmmss");
                // crude name extraction: words after 'named'
                int namedIdx = prompt.IndexOf("named ");
                if (namedIdx >= 0)
                {
                    var rest = prompt.Substring(namedIdx + 6).Trim();
                    if (!string.IsNullOrWhiteSpace(rest)) name = CultureTitle(rest);
                }
                plan.Commands.Add(new CreateSheetCommand { Name = name });
            }

            // Clear range
            if (prompt.Contains("clear "))
            {
                // Clear column letter if present
                var mcol = System.Text.RegularExpressions.Regex.Match(prompt, @"column\s+([a-zA-Z]+)");
                int c0 = context.StartCol; int r0 = context.StartRow; int rows = Math.Max(1, context.Rows); int cols = Math.Max(1, context.Cols);
                if (mcol.Success)
                {
                    try { c0 = SpreadsheetApp.Core.CellAddress.ColumnNameToIndex(mcol.Groups[1].Value); cols = 1; } catch { }
                }
                plan.Commands.Add(new ClearRangeCommand { StartRow = r0, StartCol = c0, Rows = rows, Cols = cols });
                return Task.FromResult(plan);
            }

            // Rename sheet
            if (prompt.Contains("rename sheet"))
            {
                string newName = "Sheet";
                var toIdx = prompt.IndexOf(" to ");
                if (toIdx >= 0) newName = CultureTitle(prompt.Substring(toIdx + 4).Trim());
                plan.Commands.Add(new RenameSheetCommand { NewName = newName });
                return Task.FromResult(plan);
            }

            // Sort range (use selection shape)
            if (prompt.Contains("sort"))
            {
                int r0 = context.StartRow; int c0 = context.StartCol;
                int rows = Math.Max(1, context.Rows); int cols = Math.Max(1, context.Cols);
                bool desc = prompt.Contains("desc");
                bool hasHeader = true;
                int sortCol = c0;
                // Try to locate header name in selection values
                try
                {
                    if (context.SelectionValues != null && context.SelectionValues.Length > 0)
                    {
                        var headerRow = context.SelectionValues[0];
                        for (int i = 0; i < headerRow.Length; i++)
                        {
                            var h = headerRow[i]?.Trim();
                            if (string.IsNullOrEmpty(h)) continue;
                            if (h.Equals("amount", StringComparison.OrdinalIgnoreCase)) { sortCol = c0 + i; break; }
                        }
                    }
                }
                catch { }
                // Or parse explicit column letter
                var mcol = System.Text.RegularExpressions.Regex.Match(prompt, @"column\s+([a-zA-Z]+)");
                if (mcol.Success) { try { sortCol = SpreadsheetApp.Core.CellAddress.ColumnNameToIndex(mcol.Groups[1].Value); } catch { } }
                plan.Commands.Add(new SortRangeCommand { StartRow = r0, StartCol = c0, Rows = rows, Cols = cols, SortCol = sortCol, Order = desc ? "desc" : "asc", HasHeader = hasHeader });
                return Task.FromResult(plan);
            }

            // SUM/AVERAGE formulas
            if (prompt.Contains("sum") || prompt.Contains("average"))
            {
                var ranges = System.Text.RegularExpressions.Regex.Matches(prompt, @"([A-Za-z]+\d+:[A-Za-z]+\d+)");
                var cells = System.Text.RegularExpressions.Regex.Matches(prompt, @"\b([A-Za-z]+\d+)\b");
                string? sumDest = null; string? sumRange = null; string? avgDest = null; string? avgRange = null;
                if (ranges.Count > 0) { sumRange = ranges[0].Groups[1].Value; if (ranges.Count > 1) avgRange = ranges[1].Groups[1].Value; }
                // Heuristic: first standalone cell is sum dest, next standalone cell is average dest
                var standalones = new List<string>();
                foreach (System.Text.RegularExpressions.Match m in cells) if (!m.Value.Contains(':')) standalones.Add(m.Groups[1].Value);
                if (standalones.Count > 0) sumDest = standalones[0];
                if (standalones.Count > 1) avgDest = standalones[1];
                var sf = new SetFormulaCommand { StartRow = 0, StartCol = 0, Formulas = new string[0][] };
                var rowsList = new List<string[]>();
                if (!string.IsNullOrWhiteSpace(sumDest) && !string.IsNullOrWhiteSpace(sumRange))
                {
                    if (SpreadsheetApp.Core.CellAddress.TryParse(sumDest, out int r, out int c))
                    {
                        sf.StartRow = r; sf.StartCol = c; rowsList.Add(new[] { $"=SUM({sumRange})" });
                    }
                }
                if (!string.IsNullOrWhiteSpace(avgDest) && !string.IsNullOrWhiteSpace(avgRange))
                {
                    if (rowsList.Count == 0)
                    {
                        if (SpreadsheetApp.Core.CellAddress.TryParse(avgDest, out int r2, out int c2)) { sf.StartRow = r2; sf.StartCol = c2; rowsList.Add(new[] { $"=AVERAGE({avgRange})" }); }
                    }
                    else
                    {
                        // Append below the previous formula destination
                        rowsList.Add(new[] { $"=AVERAGE({avgRange})" });
                    }
                }
                if (rowsList.Count > 0)
                {
                    sf.Formulas = rowsList.ToArray();
                    plan.Commands.Add(sf);
                    return Task.FromResult(plan);
                }
            }

            // Expense table (Date, Description, Amount) with sample rows
            if (prompt.Contains("expense") && prompt.Contains("date") && prompt.Contains("description") && prompt.Contains("amount"))
            {
                int r0 = context.StartRow; int c0 = context.StartCol;
                var vals = new List<string[]>();
                // headers
                vals.Add(new[] { "Date", "Description", "Amount" });
                // rows (3)
                vals.Add(new[] { "2023-10-01", "Groceries", "45" });
                vals.Add(new[] { "2023-10-02", "Utilities", "30" });
                vals.Add(new[] { "2023-10-03", "Supplies", "25" });
                plan.Commands.Add(new SetValuesCommand { StartRow = r0, StartCol = c0, Values = vals.ToArray() });
                return Task.FromResult(plan);
            }

            // Add Tax column at D with 10% of Amount
            if (prompt.Contains("tax") && prompt.Contains("10%"))
            {
                int r0 = context.StartRow; int c0 = context.StartCol; // assume anchored at D3:D6 per spec
                // header
                plan.Commands.Add(new SetValuesCommand { StartRow = r0, StartCol = c0, Values = new[] { new[] { "Tax" } } });
                // formulas for rows beneath header (best-effort: 3 rows)
                int dataRows = Math.Max(1, context.Rows - 1);
                var f = new string[dataRows][];
                for (int i = 0; i < dataRows; i++)
                {
                    int rr1 = r0 + 1 + i; // 0-based
                    // Amount column is immediately to the left (C)
                    string addr = SpreadsheetApp.Core.CellAddress.ToAddress(rr1, Math.Max(0, c0 - 1));
                    f[i] = new[] { $"={addr}*0.1" };
                }
                plan.Commands.Add(new SetFormulaCommand { StartRow = r0 + 1, StartCol = c0, Formulas = f });
                return Task.FromResult(plan);
            }

            // Add total row summing Amount and Tax columns
            if (prompt.Contains("total row") || prompt.Contains("sum"))
            {
                // Best-effort: assume headers at row r0, data rows next two rows
                int r0 = context.StartRow; int c0 = context.StartCol;
                int totalsRow = r0; // if location anchors at total row, use that
                // write formulas in C and D if available
                var rows = new List<string[]>();
                rows.Add(new[] { "=SUM(C4:C5)" });
                var cmdC = new SetFormulaCommand { StartRow = Math.Max(0, totalsRow), StartCol = 2, Formulas = new[] { new[] { "=SUM(C4:C5)" } } };
                plan.Commands.Add(cmdC);
                var cmdD = new SetFormulaCommand { StartRow = Math.Max(0, totalsRow), StartCol = 3, Formulas = new[] { new[] { "=SUM(D4:D5)" } } };
                plan.Commands.Add(cmdD);
                return Task.FromResult(plan);
            }

            // Bonus column = 15% of Salary
            if (prompt.Contains("bonus") && prompt.Contains("15%"))
            {
                int r0 = context.StartRow; int c0 = context.StartCol;
                // header
                plan.Commands.Add(new SetValuesCommand { StartRow = r0, StartCol = c0, Values = new[] { new[] { "Bonus" } } });
                int dataRows = Math.Max(1, context.Rows - 1);
                var f = new string[dataRows][];
                for (int i = 0; i < dataRows; i++)
                {
                    int rr1 = r0 + 1 + i;
                    // Assume Salary in column C (left of D)
                    string addr = SpreadsheetApp.Core.CellAddress.ToAddress(rr1, Math.Max(0, c0 - 1));
                    f[i] = new[] { $"={addr}*0.15" };
                }
                plan.Commands.Add(new SetFormulaCommand { StartRow = r0 + 1, StartCol = c0, Formulas = f });
                return Task.FromResult(plan);
            }

            // Fallback: set a non-invasive title above
            if (!string.IsNullOrEmpty(prompt))
            {
                int tr = Math.Max(0, context.StartRow - 1);
                int tc = context.StartCol;
                plan.Commands.Add(new SetTitleCommand { StartRow = tr, StartCol = tc, Rows = 1, Cols = 1, Text = CultureTitle(prompt) });
            }
            return Task.FromResult(plan);
        }

        private static string CultureTitle(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return string.Empty;
            return char.ToUpperInvariant(s[0]) + (s.Length > 1 ? s.Substring(1) : string.Empty);
        }
    }
}
