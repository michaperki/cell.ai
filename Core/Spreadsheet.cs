using System;
using System.Collections.Generic;
using System.Globalization;

namespace SpreadsheetApp.Core
{
    public class Spreadsheet
    {
        public const int DefaultRows = 100;
        public const int DefaultCols = 26; // A..Z

        public int Rows { get; private set; }
        public int Columns { get; private set; }

        // Store only non-empty raw values
        private readonly Dictionary<(int r, int c), string> _raw = new();

        // Last computed values for display
        private Dictionary<(int r, int c), EvaluationResult> _cache = new();

        // Dependency tracking
        // _dependencies[cell] = set of cells this cell references
        private readonly Dictionary<(int r, int c), HashSet<(int r, int c)>> _dependencies = new();
        // _dependents[cell] = set of cells that reference this cell
        private readonly Dictionary<(int r, int c), HashSet<(int r, int c)>> _dependents = new();
        
        // Cell formatting
        private readonly Dictionary<(int r, int c), CellFormat> _formats = new();

        // Optional resolver for cross-sheet references (Sheet name, row, col -> EvaluationResult)
        public Func<string, int, int, EvaluationResult>? CrossSheetResolver { get; set; }

        public Spreadsheet(int rows, int columns)
        {
            if (rows <= 0 || columns <= 0) throw new ArgumentOutOfRangeException();
            Rows = rows; Columns = columns;
        }

        public string? GetRaw(int row, int col)
        {
            _raw.TryGetValue((row, col), out var v);
            return v;
        }

        public void SetRaw(int row, int col, string? raw)
        {
            raw ??= string.Empty;
            var key = (row, col);

            // Update raw storage
            if (string.IsNullOrWhiteSpace(raw)) _raw.Remove(key);
            else _raw[key] = raw;

            // Update dependency graph for this cell
            // Remove old edges
            if (_dependencies.TryGetValue(key, out var oldRefs))
            {
                foreach (var dep in oldRefs)
                {
                    if (_dependents.TryGetValue(dep, out var rev))
                    {
                        rev.Remove(key);
                        if (rev.Count == 0) _dependents.Remove(dep);
                    }
                }
            }

            var newRefs = new HashSet<(int r, int c)>();
            if (!string.IsNullOrWhiteSpace(raw) && raw.StartsWith("=", StringComparison.Ordinal))
            {
                foreach (var rc in ExtractReferences(raw.Substring(1))) newRefs.Add(rc);
            }
            if (newRefs.Count > 0) _dependencies[key] = newRefs; else _dependencies.Remove(key);
            foreach (var dep in newRefs)
            {
                if (!_dependents.TryGetValue(dep, out var rev))
                {
                    rev = new HashSet<(int, int)>();
                    _dependents[dep] = rev;
                }
                rev.Add(key);
            }
        }

        public EvaluationResult GetValue(int row, int col)
        {
            if (_cache.TryGetValue((row, col), out var v)) return v;
            throw new InvalidOperationException("Call Recalculate() first");
        }

        public string GetDisplay(int row, int col)
        {
            return GetValue(row, col).ToDisplay();
        }

        public CellFormat? GetFormat(int row, int col)
        {
            _formats.TryGetValue((row, col), out var f);
            return f;
        }

        public void SetFormat(int row, int col, CellFormat? format)
        {
            var key = (row, col);
            if (format == null || format.IsDefault()) _formats.Remove(key);
            else _formats[key] = format;
        }

        public IEnumerable<KeyValuePair<(int r, int c), CellFormat>> GetAllFormats() => _formats;

        public void Recalculate()
        {
            _cache = new();
            var visiting = new HashSet<(int, int)>();
            for (int r = 0; r < Rows; r++)
            {
                for (int c = 0; c < Columns; c++)
                {
                    EvaluateCell(r, c, visiting);
                }
            }
        }

        public IReadOnlyCollection<(int r, int c)> RecalculateDirty(int row, int col)
        {
            var start = (row, col);
            var affected = new HashSet<(int r, int c)> { start };
            var q = new Queue<(int r, int c)>();
            q.Enqueue(start);
            while (q.Count > 0)
            {
                var cur = q.Dequeue();
                if (_dependents.TryGetValue(cur, out var rev))
                {
                    foreach (var dep in rev)
                    {
                        if (affected.Add(dep)) q.Enqueue(dep);
                    }
                }
            }

            // Invalidate cache for affected cells
            foreach (var cell in affected) _cache.Remove(cell);

            // Recompute affected cells; dependencies will be computed recursively as needed
            foreach (var cell in affected)
            {
                EvaluateCell(cell.r, cell.c, new HashSet<(int, int)>());
            }

            return affected;
        }

        private EvaluationResult EvaluateCell(int row, int col, HashSet<(int,int)> visiting)
        {
            if (_cache.TryGetValue((row, col), out var cached)) return cached;

            if (!visiting.Add((row, col)))
            {
                var cyc = EvaluationResult.FromError("CYCLE");
                _cache[(row, col)] = cyc;
                return cyc;
            }

            string raw = GetRaw(row, col) ?? string.Empty;
            EvaluationResult result;
            try
            {
                if (string.IsNullOrWhiteSpace(raw))
                {
                    result = EvaluationResult.FromText(string.Empty);
                }
                else if (raw.StartsWith("=", StringComparison.Ordinal))
                {
                    var expr = raw.Substring(1);
                    Func<string, string, EvaluationResult>? crossSheet = null;
                    if (CrossSheetResolver != null)
                    {
                        var resolver = CrossSheetResolver;
                        crossSheet = (sheet, addr) =>
                        {
                            if (!CellAddress.TryParse(addr, out int rr2, out int cc2))
                                return EvaluationResult.FromError($"Bad ref '{addr}'");
                            return resolver(sheet, rr2, cc2);
                        };
                    }
                    var engine = new FormulaEngine((addr) =>
                    {
                        if (!CellAddress.TryParse(addr, out int rr, out int cc))
                            return EvaluationResult.FromError($"Bad ref '{addr}'");
                        if (rr < 0 || rr >= Rows || cc < 0 || cc >= Columns)
                            return EvaluationResult.FromError("REF");
                        return EvaluateCell(rr, cc, visiting);
                    }, crossSheet);
                    result = engine.Evaluate(expr);
                }
                else
                {
                    if (double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double d))
                        result = EvaluationResult.FromNumber(d);
                    else
                        result = EvaluationResult.FromText(raw);
                }
            }
            catch (Exception ex)
            {
                result = EvaluationResult.FromError(ex.Message);
            }

            _cache[(row, col)] = result;
            visiting.Remove((row, col));
            return result;
        }

        // --- Helpers ---
        private static IEnumerable<(int r, int c)> ExtractReferences(string expr)
        {
            // Delegate to the formula parser for robust reference extraction.
            return FormulaEngine.EnumerateReferences(expr);
        }

        private static bool TryParseCellToken(string s, ref int i, out int row, out int col)
        {
            row = col = 0;
            int pos = i;
            // Optional $ before letters
            if (pos < s.Length && s[pos] == '$') pos++;
            int lettersStart = pos;
            while (pos < s.Length && char.IsLetter(s[pos])) pos++;
            if (pos == lettersStart) return false;
            // Optional $ before digits
            if (pos < s.Length && s[pos] == '$') pos++;
            int digitsStart = pos;
            while (pos < s.Length && char.IsDigit(s[pos])) pos++;
            if (pos == digitsStart) return false;
            string letters = s.Substring(lettersStart, pos - lettersStart).Replace("$", string.Empty);
            // digits are immediately before 'pos' but may include the optional $ we already handled; so just slice
            string digits = s.Substring(digitsStart, pos - digitsStart);
            string addr = (letters + digits).ToUpperInvariant();
            if (!CellAddress.TryParse(addr, out row, out col)) return false;
            i = pos;
            return true;
        }

        private static IEnumerable<(int r, int c)> EnumerateRangeCoords(int r1, int c1, int r2, int c2)
        {
            int rStart = Math.Min(r1, r2), rEnd = Math.Max(r1, r2);
            int cStart = Math.Min(c1, c2), cEnd = Math.Max(c1, c2);
            for (int r = rStart; r <= rEnd; r++)
                for (int c = cStart; c <= cEnd; c++)
                    yield return (r, c);
        }
    }
}
