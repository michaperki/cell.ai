using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using SpreadsheetApp.Core;

namespace SpreadsheetApp.Core.AI
{
    // Minimal host-controlled observe→reason→act loop.
    public static class AgentLoop
    {
        public sealed class Result
        {
            public AIPlan Plan { get; set; } = new AIPlan();
            public string[] Transcript { get; set; } = Array.Empty<string>();
        }

        public static async Task<Result> RunAsync(IChatPlanner planner, Spreadsheet sheet, AIContext baseContext, string userPrompt, CancellationToken ct)
        {
            // 1) Observe selection summary and column profile
            int sr = Math.Max(0, baseContext.StartRow);
            int sc = Math.Max(0, baseContext.StartCol);
            int rows = Math.Max(1, baseContext.Rows);
            int cols = Math.Max(1, baseContext.Cols);
            var transcript = new List<string>();

            var sel = ObservationTools.SummarizeSelection(sheet, sr, sc, rows, cols);
            transcript.Add($"Selection {sel.TopLeft} · {rows}x{cols} | non-empty rows: {sel.NonEmptyRowCount}");
            if (sel.HeaderRow != null && sel.HeaderRow.Length > 0)
            {
                // Show up to first 4 headers
                var hdr = string.Join(", ", sel.HeaderRow.Length > 4 ? sel.HeaderRow[..4] : sel.HeaderRow);
                transcript.Add($"Header: [{hdr}]");
            }

            // Target primary column = first column in selection
            int dataStartRow = sr;
            // If the cell above has a header, treat sr as first data row and show header separately
            try { if (sr > 0) { var above = sheet.GetRaw(sr - 1, sc); if (!string.IsNullOrWhiteSpace(above)) dataStartRow = sr; } }
            catch { }

            var profile = ObservationTools.ProfileColumn(sheet, sc, dataStartRow, rows, sel.HeaderRow != null && sel.HeaderRow.Length > 0 ? sel.HeaderRow[0] : null);
            var sb = new StringBuilder();
            sb.Append($"Col {profile.ColumnLetter} '{profile.Header ?? ""}': non-empty={profile.NonEmpty}, empty={profile.Empties}");
            if (profile.MostlyNumeric)
            {
                sb.Append($", numeric min={profile.Min?.ToString() ?? "-"} max={profile.Max?.ToString() ?? "-"} mean={profile.Mean?.ToString("0.###") ?? "-"}");
            }
            else
            {
                if (profile.TopExamples.Length > 0) sb.Append($", examples: {string.Join(" | ", profile.TopExamples)}");
            }
            transcript.Add(sb.ToString());

            // 2) Unique values for the primary column
            var uniques = ObservationTools.UniqueValues(sheet, sc, dataStartRow, rows, topK: 10);
            var uniqSb = new StringBuilder();
            uniqSb.Append($"Uniques {uniques.ColumnLetter}: total={uniques.TotalUniques}, blanks={uniques.Blanks}");
            if (uniques.Top.Count > 0)
            {
                uniqSb.Append(" | top: ");
                int k = 0;
                foreach (var (val, count) in uniques.Top)
                {
                    if (k++ > 0) uniqSb.Append(", ");
                    string v = val.Length > 24 ? val.Substring(0, 24) + "…" : val;
                    uniqSb.Append($"'{v}'×{count}");
                }
            }
            transcript.Add(uniqSb.ToString());

            // 3) Sample first few data rows in the selection (up to 5)
            int sampleRows = Math.Min(5, rows);
            var sample = ObservationTools.GetRange(sheet, sr, sc, sampleRows, Math.Min(cols, 6));
            var sampleLines = new List<string>();
            for (int r = 0; r < sample.Length; r++)
            {
                var line = string.Join(" | ", sample[r]);
                sampleLines.Add($"Row {sr + r + 1}: {line}");
            }
            transcript.AddRange(sampleLines);

            // Build an augmented user prompt including observations
            var augUser = new StringBuilder();
            augUser.AppendLine(userPrompt.Trim());
            augUser.AppendLine();
            augUser.AppendLine("Context:");
            foreach (var t in transcript) augUser.AppendLine("- " + t);

            // Keep conversation history; otherwise same context
            var ctx = new AIContext
            {
                SheetName = baseContext.SheetName,
                StartRow = baseContext.StartRow,
                StartCol = baseContext.StartCol,
                Rows = baseContext.Rows,
                Cols = baseContext.Cols,
                Title = baseContext.Title,
                SelectionValues = baseContext.SelectionValues,
                NearbyValues = baseContext.NearbyValues,
                Workbook = baseContext.Workbook,
                Conversation = baseContext.Conversation,
                AllowedCommands = baseContext.AllowedCommands,
                WritePolicy = baseContext.WritePolicy,
                Schema = baseContext.Schema
            };

            var plan = await planner.PlanAsync(ctx, augUser.ToString(), ct).ConfigureAwait(false);
            return new Result { Plan = plan, Transcript = transcript.ToArray() };
        }
    }
}
