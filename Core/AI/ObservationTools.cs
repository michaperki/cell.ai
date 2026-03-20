using System;
using System.Collections.Generic;
using System.Globalization;
using SpreadsheetApp.Core;

namespace SpreadsheetApp.Core.AI
{
    // Lightweight, zero-side-effect read utilities for the agent loop.
    public static class ObservationTools
    {
        public sealed class SelectionSummary
        {
            public string TopLeft { get; set; } = "A1";
            public int Rows { get; set; }
            public int Cols { get; set; }
            public string[]? HeaderRow { get; set; }
            public int NonEmptyRowCount { get; set; }
        }

        public sealed class ColumnProfile
        {
            public int ColumnIndex { get; set; }
            public string ColumnLetter { get; set; } = "A";
            public string? Header { get; set; }
            public int NonEmpty { get; set; }
            public int Empties { get; set; }
            public bool MostlyNumeric { get; set; }
            public double? Min { get; set; }
            public double? Max { get; set; }
            public double? Mean { get; set; }
            public string[] TopExamples { get; set; } = Array.Empty<string>();
        }

        public sealed class UniqueSummary
        {
            public int ColumnIndex { get; set; }
            public string ColumnLetter { get; set; } = "A";
            public List<(string Value, int Count)> Top { get; set; } = new();
            public int TotalUniques { get; set; }
            public int Blanks { get; set; }
        }

        public static SelectionSummary SummarizeSelection(Spreadsheet sheet, int startRow, int startCol, int rows, int cols)
        {
            var sum = new SelectionSummary
            {
                TopLeft = CellAddress.ToAddress(startRow, startCol),
                Rows = Math.Max(1, rows),
                Cols = Math.Max(1, cols)
            };
            // Header row above selection
            try
            {
                if (startRow > 0 && cols > 0)
                {
                    var hdr = new string[cols];
                    for (int c = 0; c < cols; c++) hdr[c] = sheet.GetDisplay(startRow - 1, startCol + c) ?? string.Empty;
                    sum.HeaderRow = hdr;
                }
            }
            catch { }
            // Non-empty rows count within selection (any non-empty cell across the width)
            try
            {
                int nonEmptyRows = 0;
                for (int r = 0; r < rows; r++)
                {
                    bool any = false;
                    for (int c = 0; c < cols; c++)
                    {
                        var raw = sheet.GetRaw(startRow + r, startCol + c);
                        if (!string.IsNullOrWhiteSpace(raw)) { any = true; break; }
                    }
                    if (any) nonEmptyRows++;
                }
                sum.NonEmptyRowCount = nonEmptyRows;
            }
            catch { }
            return sum;
        }

        public static ColumnProfile ProfileColumn(Spreadsheet sheet, int columnIndex, int startRow, int rowCount, string? header)
        {
            int empties = 0, nonEmpty = 0;
            bool anyNumeric = false, anyText = false;
            double min = double.MaxValue, max = double.MinValue, sum = 0.0;
            int countNumeric = 0;
            var examples = new List<string>();

            int rows = Math.Max(0, rowCount);
            for (int i = 0; i < rows; i++)
            {
                string raw = sheet.GetRaw(startRow + i, columnIndex) ?? string.Empty;
                if (string.IsNullOrWhiteSpace(raw)) { empties++; continue; }
                nonEmpty++;
                if (double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double d))
                {
                    anyNumeric = true; countNumeric++;
                    if (d < min) min = d; if (d > max) max = d; sum += d;
                }
                else { anyText = true; if (examples.Count < 5) examples.Add(raw); }
            }

            double? mean = countNumeric > 0 ? sum / countNumeric : null;
            double? minN = countNumeric > 0 ? min : null;
            double? maxN = countNumeric > 0 ? max : null;
            return new ColumnProfile
            {
                ColumnIndex = columnIndex,
                ColumnLetter = CellAddress.ColumnIndexToName(columnIndex),
                Header = header,
                NonEmpty = nonEmpty,
                Empties = empties,
                MostlyNumeric = anyNumeric && !anyText,
                Min = minN,
                Max = maxN,
                Mean = mean,
                TopExamples = examples.ToArray()
            };
        }

        public static UniqueSummary UniqueValues(Spreadsheet sheet, int columnIndex, int startRow, int rowCount, int topK = 10)
        {
            var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            int blanks = 0;
            for (int i = 0; i < rowCount; i++)
            {
                string raw = sheet.GetRaw(startRow + i, columnIndex) ?? string.Empty;
                string v = raw.Trim();
                if (string.IsNullOrWhiteSpace(v)) { blanks++; continue; }
                counts.TryGetValue(v, out int n); counts[v] = n + 1;
            }
            var top = new List<(string, int)>();
            foreach (var kv in counts)
                top.Add((kv.Key, kv.Value));
            top.Sort((a, b) => b.Item2.CompareTo(a.Item2));
            if (top.Count > topK) top = top.GetRange(0, topK);
            return new UniqueSummary
            {
                ColumnIndex = columnIndex,
                ColumnLetter = CellAddress.ColumnIndexToName(columnIndex),
                Top = top,
                TotalUniques = counts.Count,
                Blanks = blanks
            };
        }

        public static string[][] GetRange(Spreadsheet sheet, int startRow, int startCol, int rows, int cols)
        {
            rows = Math.Max(0, rows); cols = Math.Max(0, cols);
            var arr = new string[rows][];
            for (int r = 0; r < rows; r++)
            {
                arr[r] = new string[cols];
                for (int c = 0; c < cols; c++) arr[r][c] = sheet.GetDisplay(startRow + r, startCol + c) ?? string.Empty;
            }
            return arr;
        }
    }
}
