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
            int sr = Math.Max(0, baseContext.StartRow);
            int sc = Math.Max(0, baseContext.StartCol);
            int rows = Math.Max(1, baseContext.Rows);
            int cols = Math.Max(1, baseContext.Cols);
            var transcript = new List<string>();

            // Phase 1: ask the planner for query intents
            var qctx = new AIContext
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
                AllowedCommands = new string[0],
                WritePolicy = baseContext.WritePolicy,
                Schema = baseContext.Schema,
                RequestQueriesOnly = true
            };
            var qplan = await planner.PlanAsync(qctx, userPrompt, ct).ConfigureAwait(false);

            if (qplan.Queries.Count > 0)
            {
                foreach (var q in qplan.Queries)
                {
                    switch (q)
                    {
                        case SelectionSummaryQuery:
                        {
                            var sel = ObservationTools.SummarizeSelection(sheet, sr, sc, rows, cols);
                            transcript.Add($"Selection {sel.TopLeft} · {rows}x{cols} | non-empty rows: {sel.NonEmptyRowCount}");
                            if (sel.HeaderRow != null && sel.HeaderRow.Length > 0)
                            {
                                var hdr = string.Join(", ", sel.HeaderRow.Length > 4 ? sel.HeaderRow[..4] : sel.HeaderRow);
                                transcript.Add($"Header: [{hdr}]");
                            }
                            break;
                        }
                        case ProfileColumnQuery pc:
                        {
                            int dataStartRow = sr;
                            try { if (sr > 0) { var above = sheet.GetRaw(sr - 1, pc.ColumnIndex); if (!string.IsNullOrWhiteSpace(above)) dataStartRow = sr; } } catch { }
                            var prof = ObservationTools.ProfileColumn(sheet, pc.ColumnIndex, dataStartRow, rows, null);
                            var sb = new StringBuilder();
                            sb.Append($"Col {prof.ColumnLetter} '{prof.Header ?? ""}': non-empty={prof.NonEmpty}, empty={prof.Empties}");
                            if (prof.MostlyNumeric) sb.Append($", numeric min={prof.Min?.ToString() ?? "-"} max={prof.Max?.ToString() ?? "-"} mean={prof.Mean?.ToString("0.###") ?? "-"}");
                            else if (prof.TopExamples.Length > 0) sb.Append($", examples: {string.Join(" | ", prof.TopExamples)}");
                            transcript.Add(sb.ToString());
                            break;
                        }
                        case UniqueValuesQuery uq:
                        {
                            var uniques = ObservationTools.UniqueValues(sheet, uq.ColumnIndex, sr, rows, topK: uq.TopK > 0 ? uq.TopK : 10);
                            var sb = new StringBuilder();
                            sb.Append($"Uniques {uniques.ColumnLetter}: total={uniques.TotalUniques}, blanks={uniques.Blanks}");
                            if (uniques.Top.Count > 0)
                            {
                                sb.Append(" | top: ");
                                int k = 0; foreach (var (val, count) in uniques.Top) { if (k++ > 0) sb.Append(", "); string v = val.Length > 24 ? val.Substring(0, 24) + "…" : val; sb.Append($"'{v}'×{count}"); }
                            }
                            transcript.Add(sb.ToString());
                            break;
                        }
                        case SampleRowsQuery s:
                        {
                            int rr = Math.Min(Math.Max(1, s.Rows), rows);
                            int cc = Math.Min(Math.Max(1, s.Cols), cols);
                            var sample = ObservationTools.GetRange(sheet, sr, sc, rr, cc);
                            for (int r = 0; r < rr; r++) transcript.Add($"Row {sr + r + 1}: {string.Join(" | ", sample[r])}");
                            break;
                        }
                    }
                }
            }
            else
            {
                // Fallback minimal observations
                var sel = ObservationTools.SummarizeSelection(sheet, sr, sc, rows, cols);
                transcript.Add($"Selection {sel.TopLeft} · {rows}x{cols} | non-empty rows: {sel.NonEmptyRowCount}");
                if (sel.HeaderRow != null && sel.HeaderRow.Length > 0)
                {
                    var hdr = string.Join(", ", sel.HeaderRow.Length > 4 ? sel.HeaderRow[..4] : sel.HeaderRow);
                    transcript.Add($"Header: [{hdr}]");
                }
                var prof = ObservationTools.ProfileColumn(sheet, sc, sr, rows, sel.HeaderRow != null && sel.HeaderRow.Length > 0 ? sel.HeaderRow[0] : null);
                var sb = new StringBuilder();
                sb.Append($"Col {prof.ColumnLetter} '{prof.Header ?? ""}': non-empty={prof.NonEmpty}, empty={prof.Empties}");
                if (prof.MostlyNumeric) sb.Append($", numeric min={prof.Min?.ToString() ?? "-"} max={prof.Max?.ToString() ?? "-"} mean={prof.Mean?.ToString("0.###") ?? "-"}");
                else if (prof.TopExamples.Length > 0) sb.Append($", examples: {string.Join(" | ", prof.TopExamples)}");
                transcript.Add(sb.ToString());
                var uniques = ObservationTools.UniqueValues(sheet, sc, sr, rows, topK: 10);
                var us = new StringBuilder();
                us.Append($"Uniques {uniques.ColumnLetter}: total={uniques.TotalUniques}, blanks={uniques.Blanks}");
                if (uniques.Top.Count > 0)
                {
                    us.Append(" | top: ");
                    int k = 0; foreach (var (val, count) in uniques.Top) { if (k++ > 0) us.Append(", "); string v = val.Length > 24 ? val.Substring(0, 24) + "…" : val; us.Append($"'{v}'×{count}"); }
                }
                transcript.Add(us.ToString());
                int sampleRows = Math.Min(5, rows);
                var sample = ObservationTools.GetRange(sheet, sr, sc, sampleRows, Math.Min(cols, 6));
                for (int r = 0; r < sample.Length; r++) transcript.Add($"Row {sr + r + 1}: {string.Join(" | ", sample[r])}");
            }

            // Phase 2: build augmented prompt and ask for writes
            var augUser = new StringBuilder();
            augUser.AppendLine(userPrompt.Trim());
            augUser.AppendLine();
            augUser.AppendLine("Observations:");
            foreach (var t in transcript) augUser.AppendLine("- " + t);

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
                Schema = baseContext.Schema,
                RequestQueriesOnly = false
            };
            var plan = await planner.PlanAsync(ctx, augUser.ToString(), ct).ConfigureAwait(false);
            return new Result { Plan = plan, Transcript = transcript.ToArray() };
        }
    }
}
