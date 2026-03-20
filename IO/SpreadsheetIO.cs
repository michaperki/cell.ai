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
        public sealed class Workbook
        {
            public List<Spreadsheet> Sheets { get; } = new();
            public List<string> Names { get; } = new();
        }

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

        private class WorkbookData
        {
            public int FormatVersion { get; set; } = 1;
            public List<WorkbookSheet> Sheets { get; set; } = new();
        }

        private class WorkbookSheet
        {
            public string Name { get; set; } = "Sheet";
            public int Rows { get; set; }
            public int Columns { get; set; }
            public Dictionary<string, string> Cells { get; set; } = new();
            public Dictionary<string, FormatData>? Formats { get; set; }
        }

        public static void SaveWorkbookToFile(System.Collections.Generic.IReadOnlyList<Spreadsheet> sheets, System.Collections.Generic.IReadOnlyList<string> names, string path)
        {
            var wb = new WorkbookData { FormatVersion = 1 };
            for (int i = 0; i < sheets.Count; i++)
            {
                var s = sheets[i];
                var ws = new WorkbookSheet { Name = (i < names.Count ? names[i] : $"Sheet{i + 1}"), Rows = s.Rows, Columns = s.Columns };
                for (int r = 0; r < s.Rows; r++)
                {
                    for (int c = 0; c < s.Columns; c++)
                    {
                        var raw = s.GetRaw(r, c);
                        if (!string.IsNullOrWhiteSpace(raw)) ws.Cells[CellAddress.ToAddress(r, c)] = raw!;
                    }
                }
                var formats = new Dictionary<string, FormatData>();
                foreach (var kv in s.GetAllFormats())
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
                if (formats.Count > 0) ws.Formats = formats;
                wb.Sheets.Add(ws);
            }
            var opts = new JsonSerializerOptions { WriteIndented = true };
            var json = JsonSerializer.Serialize(wb, opts);
            File.WriteAllText(path, json);
        }

        public static Workbook LoadWorkbookFromFile(string path)
        {
            var json = File.ReadAllText(path);
            try
            {
                using var doc = JsonDocument.Parse(json);
                if (doc.RootElement.ValueKind == JsonValueKind.Object && doc.RootElement.TryGetProperty("Sheets", out var _))
                {
                    // Workbook format
                    var wbData = JsonSerializer.Deserialize<WorkbookData>(json) ?? new WorkbookData();
                    var wb = new Workbook();
                    foreach (var ws in wbData.Sheets)
                    {
                        var sheet = new Spreadsheet(ws.Rows > 0 ? ws.Rows : Spreadsheet.DefaultRows, ws.Columns > 0 ? ws.Columns : Spreadsheet.DefaultCols);
                        foreach (var kv in ws.Cells)
                        {
                            if (CellAddress.TryParse(kv.Key, out int r, out int c)) sheet.SetRaw(r, c, kv.Value);
                        }
                        if (ws.Formats != null)
                        {
                            foreach (var fk in ws.Formats)
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
                        wb.Sheets.Add(sheet);
                        wb.Names.Add(ws.Name ?? $"Sheet{wb.Names.Count + 1}");
                    }
                    return wb;
                }
                else
                {
                    // Single-sheet fallback
                    var s = LoadFromFile(path);
                    var wb = new Workbook();
                    wb.Sheets.Add(s); wb.Names.Add("Sheet1");
                    return wb;
                }
            }
            catch
            {
                // Fallback on parse errors
                var s = LoadFromFile(path);
                var wb = new Workbook();
                wb.Sheets.Add(s); wb.Names.Add("Sheet1");
                return wb;
            }
        }

        public static async Task SaveWorkbookToFileAsync(System.Collections.Generic.IReadOnlyList<Spreadsheet> sheets, System.Collections.Generic.IReadOnlyList<string> names, string path)
        {
            var wb = new WorkbookData { FormatVersion = 1 };
            for (int i = 0; i < sheets.Count; i++)
            {
                var s = sheets[i];
                var ws = new WorkbookSheet { Name = (i < names.Count ? names[i] : $"Sheet{i + 1}"), Rows = s.Rows, Columns = s.Columns };
                for (int r = 0; r < s.Rows; r++)
                {
                    for (int c = 0; c < s.Columns; c++)
                    {
                        var raw = s.GetRaw(r, c);
                        if (!string.IsNullOrWhiteSpace(raw)) ws.Cells[CellAddress.ToAddress(r, c)] = raw!;
                    }
                }
                var formats = new Dictionary<string, FormatData>();
                foreach (var kv in s.GetAllFormats())
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
                if (formats.Count > 0) ws.Formats = formats;
                wb.Sheets.Add(ws);
            }
            var opts = new JsonSerializerOptions { WriteIndented = true };
            var json = JsonSerializer.Serialize(wb, opts);
            await File.WriteAllTextAsync(path, json);
        }

        public static async Task<Workbook> LoadWorkbookFromFileAsync(string path)
        {
            var json = await File.ReadAllTextAsync(path);
            try
            {
                using var doc = JsonDocument.Parse(json);
                if (doc.RootElement.ValueKind == JsonValueKind.Object && doc.RootElement.TryGetProperty("Sheets", out var _))
                {
                    var wbData = JsonSerializer.Deserialize<WorkbookData>(json) ?? new WorkbookData();
                    var wb = new Workbook();
                    foreach (var ws in wbData.Sheets)
                    {
                        var sheet = new Spreadsheet(ws.Rows > 0 ? ws.Rows : Spreadsheet.DefaultRows, ws.Columns > 0 ? ws.Columns : Spreadsheet.DefaultCols);
                        foreach (var kv in ws.Cells)
                        {
                            if (CellAddress.TryParse(kv.Key, out int r, out int c)) sheet.SetRaw(r, c, kv.Value);
                        }
                        if (ws.Formats != null)
                        {
                            foreach (var fk in ws.Formats)
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
                        wb.Sheets.Add(sheet);
                        wb.Names.Add(ws.Name ?? $"Sheet{wb.Names.Count + 1}");
                    }
                    return wb;
                }
                else
                {
                    // Single-sheet fallback
                    var s = await LoadFromFileAsync(path);
                    var wb = new Workbook();
                    wb.Sheets.Add(s); wb.Names.Add("Sheet1");
                    return wb;
                }
            }
            catch
            {
                var s = await LoadFromFileAsync(path);
                var wb = new Workbook();
                wb.Sheets.Add(s); wb.Names.Add("Sheet1");
                return wb;
            }
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

        public static async Task ExportCsvAsync(Spreadsheet sheet, string path)
        {
            await using var sw = new StreamWriter(path);
            for (int r = 0; r < sheet.Rows; r++)
            {
                var fields = new List<string>(sheet.Columns);
                for (int c = 0; c < sheet.Columns; c++)
                {
                    string v = sheet.GetDisplay(r, c) ?? string.Empty;
                    fields.Add(CsvEscape(v));
                }
                await sw.WriteLineAsync(string.Join(',', fields));
            }
        }

        public static async Task<Spreadsheet> ImportCsvAsync(string path)
        {
            var lines = await File.ReadAllLinesAsync(path);
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
