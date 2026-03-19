using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using SpreadsheetApp.Core;
using System.Threading.Tasks;

namespace SpreadsheetApp.IO
{
    public static class SpreadsheetIO
    {
        private class SheetData
        {
            public int Rows { get; set; }
            public int Columns { get; set; }
            public Dictionary<string, string> Cells { get; set; } = new();
            public Dictionary<string, FormatData>? Formats { get; set; }
        }

        private class FormatData
        {
            public bool Bold { get; set; }
            public int? ForeColorArgb { get; set; }
            public int? BackColorArgb { get; set; }
            public string? NumberFormat { get; set; }
            public string? HAlign { get; set; }
        }

        public static void SaveToFile(Spreadsheet sheet, string path)
        {
            var data = new SheetData { Rows = sheet.Rows, Columns = sheet.Columns };
            for (int r = 0; r < sheet.Rows; r++)
            {
                for (int c = 0; c < sheet.Columns; c++)
                {
                    var raw = sheet.GetRaw(r, c);
                    if (!string.IsNullOrWhiteSpace(raw))
                    {
                        data.Cells[CellAddress.ToAddress(r, c)] = raw!;
                    }
                }
            }

            // Formats
            var formats = new Dictionary<string, FormatData>();
            foreach (var kv in sheet.GetAllFormats())
            {
                string addr = CellAddress.ToAddress(kv.Key.r, kv.Key.c);
                var f = kv.Value;
                formats[addr] = new FormatData
                {
                    Bold = f.Bold,
                    ForeColorArgb = f.ForeColorArgb,
                    BackColorArgb = f.BackColorArgb,
                    NumberFormat = f.NumberFormat,
                    HAlign = f.HAlign.ToString()
                };
            }
            if (formats.Count > 0) data.Formats = formats;

            var opts = new JsonSerializerOptions { WriteIndented = true };
            var json = JsonSerializer.Serialize(data, opts);
            File.WriteAllText(path, json);
        }

        public static Spreadsheet LoadFromFile(string path)
        {
            var json = File.ReadAllText(path);
            var data = JsonSerializer.Deserialize<SheetData>(json) ?? new SheetData { Rows = Spreadsheet.DefaultRows, Columns = Spreadsheet.DefaultCols };
            var sheet = new Spreadsheet(
                data.Rows > 0 ? data.Rows : Spreadsheet.DefaultRows,
                data.Columns > 0 ? data.Columns : Spreadsheet.DefaultCols);
            foreach (var kv in data.Cells)
            {
                if (CellAddress.TryParse(kv.Key, out int r, out int c))
                {
                    sheet.SetRaw(r, c, kv.Value);
                }
            }

            if (data.Formats != null)
            {
                foreach (var fk in data.Formats)
                {
                    if (CellAddress.TryParse(fk.Key, out int r, out int c))
                    {
                        var f = fk.Value;
                        var fmt = new Core.CellFormat
                        {
                            Bold = f.Bold,
                            ForeColorArgb = f.ForeColorArgb,
                            BackColorArgb = f.BackColorArgb,
                            NumberFormat = f.NumberFormat,
                            HAlign = Enum.TryParse<Core.CellHAlign>(f.HAlign ?? "Left", out var ha) ? ha : Core.CellHAlign.Left
                        };
                        sheet.SetFormat(r, c, fmt);
                    }
                }
            }
            sheet.Recalculate();
            return sheet;
        }

        public static async Task SaveToFileAsync(Spreadsheet sheet, string path)
        {
            // mirror SaveToFile but async
            var data = new SheetData { Rows = sheet.Rows, Columns = sheet.Columns };
            for (int r = 0; r < sheet.Rows; r++)
            {
                for (int c = 0; c < sheet.Columns; c++)
                {
                    var raw = sheet.GetRaw(r, c);
                    if (!string.IsNullOrWhiteSpace(raw))
                    {
                        data.Cells[CellAddress.ToAddress(r, c)] = raw!;
                    }
                }
            }
            var formats = new Dictionary<string, FormatData>();
            foreach (var kv in sheet.GetAllFormats())
            {
                string addr = CellAddress.ToAddress(kv.Key.r, kv.Key.c);
                var f = kv.Value;
                formats[addr] = new FormatData
                {
                    Bold = f.Bold,
                    ForeColorArgb = f.ForeColorArgb,
                    BackColorArgb = f.BackColorArgb,
                    NumberFormat = f.NumberFormat,
                    HAlign = f.HAlign.ToString()
                };
            }
            if (formats.Count > 0) data.Formats = formats;
            var opts = new JsonSerializerOptions { WriteIndented = true };
            var json = JsonSerializer.Serialize(data, opts);
            await File.WriteAllTextAsync(path, json);
        }

        public static async Task<Spreadsheet> LoadFromFileAsync(string path)
        {
            var json = await File.ReadAllTextAsync(path);
            var data = JsonSerializer.Deserialize<SheetData>(json) ?? new SheetData { Rows = Spreadsheet.DefaultRows, Columns = Spreadsheet.DefaultCols };
            var sheet = new Spreadsheet(
                data.Rows > 0 ? data.Rows : Spreadsheet.DefaultRows,
                data.Columns > 0 ? data.Columns : Spreadsheet.DefaultCols);
            foreach (var kv in data.Cells)
            {
                if (CellAddress.TryParse(kv.Key, out int r, out int c))
                {
                    sheet.SetRaw(r, c, kv.Value);
                }
            }
            if (data.Formats != null)
            {
                foreach (var fk in data.Formats)
                {
                    if (CellAddress.TryParse(fk.Key, out int r, out int c))
                    {
                        var f = fk.Value;
                        var fmt = new Core.CellFormat
                        {
                            Bold = f.Bold,
                            ForeColorArgb = f.ForeColorArgb,
                            BackColorArgb = f.BackColorArgb,
                            NumberFormat = f.NumberFormat,
                            HAlign = Enum.TryParse<Core.CellHAlign>(f.HAlign ?? "Left", out var ha) ? ha : Core.CellHAlign.Left
                        };
                        sheet.SetFormat(r, c, fmt);
                    }
                }
            }
            sheet.Recalculate();
            return sheet;
        }

        // CSV export/import for a single sheet
        public static void ExportCsv(Spreadsheet sheet, string path)
        {
            using var sw = new StreamWriter(path);
            for (int r = 0; r < sheet.Rows; r++)
            {
                var fields = new List<string>(sheet.Columns);
                for (int c = 0; c < sheet.Columns; c++)
                {
                    string v = sheet.GetDisplay(r, c) ?? string.Empty;
                    fields.Add(CsvEscape(v));
                }
                sw.WriteLine(string.Join(',', fields));
            }
        }

        public static Spreadsheet ImportCsv(string path)
        {
            var lines = File.ReadAllLines(path);
            int rows = lines.Length;
            int cols = 0;
            var parsed = new List<string[]>();
            foreach (var line in lines)
            {
                var row = CsvParse(line);
                parsed.Add(row);
                cols = Math.Max(cols, row.Length);
            }
            var sheet = new Spreadsheet(rows > 0 ? rows : Spreadsheet.DefaultRows, cols > 0 ? cols : Spreadsheet.DefaultCols);
            for (int r = 0; r < parsed.Count; r++)
            {
                var row = parsed[r];
                for (int c = 0; c < row.Length; c++)
                {
                    sheet.SetRaw(r, c, row[c]);
                }
            }
            sheet.Recalculate();
            return sheet;
        }

        private static string CsvEscape(string s)
        {
            if (s.Contains('"') || s.Contains(',') || s.Contains('\n') || s.Contains('\r'))
            {
                return '"' + s.Replace("\"", "\"\"") + '"';
            }
            return s;
        }

        private static string[] CsvParse(string line)
        {
            var result = new List<string>();
            int i = 0;
            while (i < line.Length)
            {
                if (line[i] == '"')
                {
                    i++;
                    var sb = new System.Text.StringBuilder();
                    while (i < line.Length)
                    {
                        if (line[i] == '"')
                        {
                            if (i + 1 < line.Length && line[i + 1] == '"') { sb.Append('"'); i += 2; continue; }
                            i++; break;
                        }
                        sb.Append(line[i]); i++;
                    }
                    result.Add(sb.ToString());
                    if (i < line.Length && line[i] == ',') i++;
                }
                else
                {
                    int start = i;
                    while (i < line.Length && line[i] != ',') i++;
                    result.Add(line.Substring(start, i - start));
                    if (i < line.Length && line[i] == ',') i++;
                }
            }
            return result.ToArray();
        }
    }
}
