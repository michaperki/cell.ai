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

            // If prompt contains a name, set a title above selection (unless explicitly suppressed)
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
