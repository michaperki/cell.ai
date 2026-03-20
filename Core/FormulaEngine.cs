using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace SpreadsheetApp.Core
{
    public class FormulaEngine
    {
        private readonly Func<string, EvaluationResult> _cellResolver;

        public FormulaEngine(Func<string, EvaluationResult> cellResolver)
        {
            _cellResolver = cellResolver;
        }

        // Enumerate referenced cell coordinates (including expanded ranges) from an expression string.
        // This uses the parser/AST to avoid false positives in string literals and to cover nested functions.
        public static System.Collections.Generic.IEnumerable<(int r, int c)> EnumerateReferences(string expr)
        {
            var refs = new System.Collections.Generic.HashSet<(int r, int c)>();
            try
            {
                var parser = new Parser(expr ?? string.Empty);
                var node = parser.ParseExpression();
                CollectRefs(node, refs);
            }
            catch
            {
                // On parse failure, return empty to avoid incorrect graph edges
            }
            return refs;
        }

        private static void CollectRefs(Node n, System.Collections.Generic.HashSet<(int r, int c)> refs)
        {
            switch (n)
            {
                case NumberNode:
                case StringNode:
                    return;
                case CellRefNode cr:
                    if (CellAddress.TryParse(cr.Address, out int rr, out int cc)) refs.Add((rr, cc));
                    return;
                case RangeNode rn:
                    foreach (var addr in EnumerateRange(rn.A, rn.B))
                    {
                        if (CellAddress.TryParse(addr, out int r1, out int c1)) refs.Add((r1, c1));
                    }
                    return;
                case UnaryNode un:
                    CollectRefs(un.Inner, refs);
                    return;
                case BinaryNode bn:
                    CollectRefs(bn.Left, refs);
                    CollectRefs(bn.Right, refs);
                    return;
                case FuncCallNode fn:
                    foreach (var a in fn.Args) CollectRefs(a, refs);
                    return;
                default:
                    return;
            }
        }

        public EvaluationResult Evaluate(string expr)
        {
            var parser = new Parser(expr);
            var node = parser.ParseExpression();
            if (!parser.AtEnd)
            {
                return EvaluationResult.FromError("Unexpected tokens");
            }
            return EvalNode(node);
        }

        private EvaluationResult EvalNode(Node n)
        {
            try
            {
                switch (n)
                {
                    case NumberNode nn:
                        return EvaluationResult.FromNumber(nn.Value);
                    case StringNode sn:
                        return EvaluationResult.FromText(sn.Value);
                    case CellRefNode cr:
                        return _cellResolver(cr.Address);
                    case UnaryNode un:
                    {
                        var v = EvalNode(un.Inner);
                        if (v.Error != null) return v;
                        if (v.Number is double num)
                        {
                            return un.Op == TokenType.Minus ? EvaluationResult.FromNumber(-num) : EvaluationResult.FromNumber(+num);
                        }
                        return EvaluationResult.FromError("Unary op on non-number");
                    }
                    case BinaryNode bn:
                    {
                        var lv = EvalNode(bn.Left);
                        if (lv.Error != null) return lv;
                        var rv = EvalNode(bn.Right);
                        if (rv.Error != null) return rv;

                        bool IsComparison(TokenType t) => t is TokenType.Eq or TokenType.NotEq or TokenType.Lt or TokenType.Gt or TokenType.LtEq or TokenType.GtEq;

                        if (IsComparison(bn.Op))
                        {
                            if (lv.Number is not double la || rv.Number is not double rb)
                                return EvaluationResult.FromError("Comparison needs numbers");
                            bool cmp = bn.Op switch
                            {
                                TokenType.Eq => la == rb,
                                TokenType.NotEq => la != rb,
                                TokenType.Lt => la < rb,
                                TokenType.Gt => la > rb,
                                TokenType.LtEq => la <= rb,
                                TokenType.GtEq => la >= rb,
                                _ => false
                            };
                            return EvaluationResult.FromNumber(cmp ? 1.0 : 0.0);
                        }

                        if (bn.Op == TokenType.Ampersand)
                        {
                            var lstr = lv.ToDisplay();
                            var rstr = rv.ToDisplay();
                            return EvaluationResult.FromText(lstr + rstr);
                        }

                        if (lv.Number is not double a || rv.Number is not double b)
                            return EvaluationResult.FromError("Math needs numbers");
                        return bn.Op switch
                        {
                            TokenType.Plus => EvaluationResult.FromNumber(a + b),
                            TokenType.Minus => EvaluationResult.FromNumber(a - b),
                            TokenType.Star => EvaluationResult.FromNumber(a * b),
                            TokenType.Slash => b == 0
                                ? EvaluationResult.FromError("DIV/0!")
                                : EvaluationResult.FromNumber(a / b),
                            TokenType.Caret => EvaluationResult.FromNumber(Math.Pow(a, b)),
                            _ => EvaluationResult.FromError("Unknown op")
                        };
                    }
                    case FuncCallNode fn:
                        return EvalFunc(fn);
                    case RangeNode rn:
                        return EvaluationResult.FromError("Range not allowed here");
                    default:
                        return EvaluationResult.FromError("Bad expression");
                }
            }
            catch (Exception ex)
            {
                return EvaluationResult.FromError(ex.Message);
            }
        }

        private EvaluationResult EvalFunc(FuncCallNode fn)
        {
            string name = fn.Name.ToUpperInvariant();

            IEnumerable<double> AsNumbers(Node arg)
            {
                if (arg is RangeNode rng)
                {
                    foreach (var addr in EnumerateRange(rng.A, rng.B))
                    {
                        var v = _cellResolver(addr);
                        if (v.Error == null && v.Number is double d) yield return d;
                    }
                    yield break;
                }
                var ev = EvalNode(arg);
                if (ev.Error == null && ev.Number is double d1) yield return d1;
            }

            string AsText(Node arg)
            {
                if (arg is RangeNode)
                    throw new Exception("Text function does not accept range");
                var ev = EvalNode(arg);
                if (ev.Error != null) throw new Exception(ev.Error);
                return ev.ToDisplay();
            }

            bool Truthy(Node arg)
            {
                if (arg is RangeNode)
                    throw new Exception("Logical function does not accept range");
                var ev = EvalNode(arg);
                if (ev.Error != null) throw new Exception(ev.Error);
                if (ev.Number is double d) return d != 0;
                var s = ev.Text ?? string.Empty;
                return !string.IsNullOrEmpty(s);
            }

            double GetNumber(Node arg, string expected)
            {
                if (arg is RangeNode)
                    throw new Exception($"{expected} expects number");
                var ev = EvalNode(arg);
                if (ev.Error != null) throw new Exception(ev.Error);
                if (ev.Number is double d) return d;
                throw new Exception($"{expected} expects number");
            }

            switch (name)
            {
                case "IFERROR":
                {
                    if (fn.Args.Count is not 2)
                        return EvaluationResult.FromError("IFERROR expects 2 arguments");
                    var first = EvalNode(fn.Args[0]);
                    if (first.Error != null)
                    {
                        try { return EvalNode(fn.Args[1]); }
                        catch (Exception ex) { return EvaluationResult.FromError(ex.Message); }
                    }
                    return first;
                }
                case "SUM":
                {
                    if (fn.Args.Count < 1) return EvaluationResult.FromError("SUM expects at least 1 argument");
                    double s = 0;
                    foreach (var a in fn.Args.SelectMany(AsNumbers)) s += a;
                    return EvaluationResult.FromNumber(s);
                }
                case "SUMIF":
                {
                    if (fn.Args.Count is not 2 and not 3) return EvaluationResult.FromError("SUMIF expects 2 or 3 arguments");
                    if (fn.Args[0] is not RangeNode range)
                        return EvaluationResult.FromError("SUMIF first arg must be a range");
                    string crit;
                    try { crit = AsText(fn.Args[1]); }
                    catch (Exception ex) { return EvaluationResult.FromError(ex.Message); }
                    var cellAddrs = EnumerateRange(range.A, range.B).ToArray();
                    string[] sumAddrs;
                    if (fn.Args.Count == 3)
                    {
                        if (fn.Args[2] is not RangeNode sumRange)
                            return EvaluationResult.FromError("SUMIF sum_range must be a range");
                        var arr = EnumerateRange(sumRange.A, sumRange.B).ToArray();
                        if (arr.Length != cellAddrs.Length) return EvaluationResult.FromError("SUMIF ranges must be same size");
                        sumAddrs = arr;
                    }
                    else
                    {
                        sumAddrs = cellAddrs;
                    }

                    bool Match(EvaluationResult v)
                    {
                        string c = crit.Trim();
                        if (string.IsNullOrEmpty(c)) return false;
                        // Parse operator prefix
                        string op = "="; string rhs = c;
                        if (c.StartsWith("<=") || c.StartsWith(">=") || c.StartsWith("<>")) { op = c.Substring(0,2); rhs = c.Substring(2); }
                        else if (c.StartsWith("<") || c.StartsWith(">") || c.StartsWith("=")) { op = c.Substring(0,1); rhs = c.Substring(1); }
                        rhs = rhs.Trim();
                        // Try numeric compare
                        if (double.TryParse(rhs, NumberStyles.Float, CultureInfo.InvariantCulture, out double rhsNum) && v.Error == null && v.Number is double lnum)
                        {
                            return op switch
                            {
                                "=" => lnum == rhsNum,
                                "<>" => lnum != rhsNum,
                                ">" => lnum > rhsNum,
                                ">=" => lnum >= rhsNum,
                                "<" => lnum < rhsNum,
                                "<=" => lnum <= rhsNum,
                                _ => false
                            };
                        }
                        // Fall back to string equality contains only '=' or no op
                        var lstr = v.ToDisplay();
                        if (op == "=" || string.IsNullOrEmpty(op)) return string.Equals(lstr, rhs, StringComparison.OrdinalIgnoreCase);
                        if (op == "<>") return !string.Equals(lstr, rhs, StringComparison.OrdinalIgnoreCase);
                        return false;
                    }

                    double sum = 0;
                    for (int i = 0; i < cellAddrs.Length; i++)
                    {
                        var v = _cellResolver(cellAddrs[i]);
                        if (Match(v))
                        {
                            var sv = _cellResolver(sumAddrs[i]);
                            if (sv.Error == null && sv.Number is double d) sum += d;
                        }
                    }
                    return EvaluationResult.FromNumber(sum);
                }
                case "COUNTIF":
                {
                    if (fn.Args.Count != 2) return EvaluationResult.FromError("COUNTIF expects 2 arguments");
                    if (fn.Args[0] is not RangeNode range)
                        return EvaluationResult.FromError("COUNTIF first arg must be a range");
                    string crit;
                    try { crit = AsText(fn.Args[1]); }
                    catch (Exception ex) { return EvaluationResult.FromError(ex.Message); }
                    var cellAddrs = EnumerateRange(range.A, range.B).ToArray();
                    bool Match(EvaluationResult v)
                    {
                        string c = crit.Trim();
                        if (string.IsNullOrEmpty(c)) return false;
                        string op = "="; string rhs = c;
                        if (c.StartsWith("<=") || c.StartsWith(">=") || c.StartsWith("<>")) { op = c.Substring(0,2); rhs = c.Substring(2); }
                        else if (c.StartsWith("<") || c.StartsWith(">") || c.StartsWith("=")) { op = c.Substring(0,1); rhs = c.Substring(1); }
                        rhs = rhs.Trim();
                        if (double.TryParse(rhs, NumberStyles.Float, CultureInfo.InvariantCulture, out double rhsNum) && v.Error == null && v.Number is double lnum)
                        {
                            return op switch
                            {
                                "=" => lnum == rhsNum,
                                "<>" => lnum != rhsNum,
                                ">" => lnum > rhsNum,
                                ">=" => lnum >= rhsNum,
                                "<" => lnum < rhsNum,
                                "<=" => lnum <= rhsNum,
                                _ => false
                            };
                        }
                        var lstr = v.ToDisplay();
                        if (op == "=" || string.IsNullOrEmpty(op)) return string.Equals(lstr, rhs, StringComparison.OrdinalIgnoreCase);
                        if (op == "<>") return !string.Equals(lstr, rhs, StringComparison.OrdinalIgnoreCase);
                        return false;
                    }
                    int count = 0;
                    foreach (var addr in cellAddrs)
                    {
                        var v = _cellResolver(addr);
                        if (Match(v)) count++;
                    }
                    return EvaluationResult.FromNumber(count);
                }
                case "AVG":
                case "AVERAGE":
                {
                    if (fn.Args.Count < 1) return EvaluationResult.FromError("AVERAGE expects at least 1 argument");
                    var nums = fn.Args.SelectMany(AsNumbers).ToArray();
                    if (nums.Length == 0) return EvaluationResult.FromError("AVERAGE needs numbers");
                    return EvaluationResult.FromNumber(nums.Average());
                }
                case "AVERAGEIF":
                {
                    if (fn.Args.Count is not 2 and not 3) return EvaluationResult.FromError("AVERAGEIF expects 2 or 3 arguments");
                    if (fn.Args[0] is not RangeNode range)
                        return EvaluationResult.FromError("AVERAGEIF first arg must be a range");
                    string crit;
                    try { crit = AsText(fn.Args[1]); }
                    catch (Exception ex) { return EvaluationResult.FromError(ex.Message); }
                    var cellAddrs = EnumerateRange(range.A, range.B).ToArray();
                    string[] avgAddrs;
                    if (fn.Args.Count == 3)
                    {
                        if (fn.Args[2] is not RangeNode sumRange)
                            return EvaluationResult.FromError("AVERAGEIF average_range must be a range");
                        var arr = EnumerateRange(sumRange.A, sumRange.B).ToArray();
                        if (arr.Length != cellAddrs.Length) return EvaluationResult.FromError("AVERAGEIF ranges must be same size");
                        avgAddrs = arr;
                    }
                    else
                    {
                        avgAddrs = cellAddrs;
                    }

                    bool Match(EvaluationResult v)
                    {
                        string c = crit.Trim();
                        if (string.IsNullOrEmpty(c)) return false;
                        string op = "="; string rhs = c;
                        if (c.StartsWith("<=") || c.StartsWith(">=") || c.StartsWith("<>")) { op = c.Substring(0,2); rhs = c.Substring(2); }
                        else if (c.StartsWith("<") || c.StartsWith(">") || c.StartsWith("=")) { op = c.Substring(0,1); rhs = c.Substring(1); }
                        rhs = rhs.Trim();
                        if (double.TryParse(rhs, NumberStyles.Float, CultureInfo.InvariantCulture, out double rhsNum) && v.Error == null && v.Number is double lnum)
                        {
                            return op switch
                            {
                                "=" => lnum == rhsNum,
                                "<>" => lnum != rhsNum,
                                ">" => lnum > rhsNum,
                                ">=" => lnum >= rhsNum,
                                "<" => lnum < rhsNum,
                                "<=" => lnum <= rhsNum,
                                _ => false
                            };
                        }
                        var lstr = v.ToDisplay();
                        if (op == "=" || string.IsNullOrEmpty(op)) return string.Equals(lstr, rhs, StringComparison.OrdinalIgnoreCase);
                        if (op == "<>") return !string.Equals(lstr, rhs, StringComparison.OrdinalIgnoreCase);
                        return false;
                    }

                    double sum = 0; int count = 0;
                    for (int i = 0; i < cellAddrs.Length; i++)
                    {
                        var v = _cellResolver(cellAddrs[i]);
                        if (Match(v))
                        {
                            var av = _cellResolver(avgAddrs[i]);
                            if (av.Error == null && av.Number is double d) { sum += d; count++; }
                        }
                    }
                    if (count == 0) return EvaluationResult.FromError("AVERAGEIF no matching numbers");
                    return EvaluationResult.FromNumber(sum / count);
                }
                case "MIN":
                {
                    if (fn.Args.Count < 1) return EvaluationResult.FromError("MIN expects at least 1 argument");
                    var nums = fn.Args.SelectMany(AsNumbers).ToArray();
                    if (nums.Length == 0) return EvaluationResult.FromError("MIN needs numbers");
                    return EvaluationResult.FromNumber(nums.Min());
                }
                case "MAX":
                {
                    if (fn.Args.Count < 1) return EvaluationResult.FromError("MAX expects at least 1 argument");
                    var nums = fn.Args.SelectMany(AsNumbers).ToArray();
                    if (nums.Length == 0) return EvaluationResult.FromError("MAX needs numbers");
                    return EvaluationResult.FromNumber(nums.Max());
                }
                case "COUNT":
                {
                    if (fn.Args.Count < 1) return EvaluationResult.FromError("COUNT expects at least 1 argument");
                    int count = fn.Args.SelectMany(AsNumbers).Count();
                    return EvaluationResult.FromNumber(count);
                }
                case "IF":
                {
                    if (fn.Args.Count is not 2 and not 3)
                        return EvaluationResult.FromError("IF expects 2 or 3 arguments");
                    bool cond;
                    try { cond = Truthy(fn.Args[0]); }
                    catch (Exception ex) { return EvaluationResult.FromError(ex.Message); }
                    if (cond)
                    {
                        return EvalNode(fn.Args[1]);
                    }
                    else
                    {
                        if (fn.Args.Count == 3) return EvalNode(fn.Args[2]);
                        return EvaluationResult.FromText(string.Empty);
                    }
                }
                case "AND":
                {
                    if (fn.Args.Count < 1)
                        return EvaluationResult.FromError("AND expects at least 1 argument");
                    foreach (var a in fn.Args)
                    {
                        try { if (!Truthy(a)) return EvaluationResult.FromNumber(0.0); }
                        catch (Exception ex) { return EvaluationResult.FromError(ex.Message); }
                    }
                    return EvaluationResult.FromNumber(1.0);
                }
                case "OR":
                {
                    if (fn.Args.Count < 1)
                        return EvaluationResult.FromError("OR expects at least 1 argument");
                    foreach (var a in fn.Args)
                    {
                        try { if (Truthy(a)) return EvaluationResult.FromNumber(1.0); }
                        catch (Exception ex) { return EvaluationResult.FromError(ex.Message); }
                    }
                    return EvaluationResult.FromNumber(0.0);
                }
                case "NOT":
                {
                    if (fn.Args.Count != 1)
                        return EvaluationResult.FromError("NOT expects 1 argument");
                    try { return EvaluationResult.FromNumber(Truthy(fn.Args[0]) ? 0.0 : 1.0); }
                    catch (Exception ex) { return EvaluationResult.FromError(ex.Message); }
                }
                case "LEN":
                {
                    if (fn.Args.Count != 1) return EvaluationResult.FromError("LEN expects 1 argument");
                    try { var s = AsText(fn.Args[0]); return EvaluationResult.FromNumber(s.Length); }
                    catch (Exception ex) { return EvaluationResult.FromError(ex.Message); }
                }
                case "LEFT":
                {
                    if (fn.Args.Count != 2) return EvaluationResult.FromError("LEFT expects 2 arguments");
                    try
                    {
                        var s = AsText(fn.Args[0]);
                        var n = (int)Math.Max(0, Math.Floor(GetNumber(fn.Args[1], "LEFT")));
                        if (n >= s.Length) return EvaluationResult.FromText(s);
                        return EvaluationResult.FromText(s.Substring(0, n));
                    }
                    catch (Exception ex) { return EvaluationResult.FromError(ex.Message); }
                }
                case "RIGHT":
                {
                    if (fn.Args.Count != 2) return EvaluationResult.FromError("RIGHT expects 2 arguments");
                    try
                    {
                        var s = AsText(fn.Args[0]);
                        var n = (int)Math.Max(0, Math.Floor(GetNumber(fn.Args[1], "RIGHT")));
                        if (n >= s.Length) return EvaluationResult.FromText(s);
                        return EvaluationResult.FromText(s.Substring(s.Length - n, n));
                    }
                    catch (Exception ex) { return EvaluationResult.FromError(ex.Message); }
                }
                case "MID":
                {
                    if (fn.Args.Count != 3) return EvaluationResult.FromError("MID expects 3 arguments");
                    try
                    {
                        var s = AsText(fn.Args[0]);
                        int start = (int)Math.Floor(GetNumber(fn.Args[1], "MID")); // 1-based
                        int len = (int)Math.Max(0, Math.Floor(GetNumber(fn.Args[2], "MID")));
                        if (start < 1) start = 1;
                        int zeroBased = start - 1;
                        if (zeroBased >= s.Length || len == 0) return EvaluationResult.FromText(string.Empty);
                        int take = Math.Min(len, s.Length - zeroBased);
                        return EvaluationResult.FromText(s.Substring(zeroBased, take));
                    }
                    catch (Exception ex) { return EvaluationResult.FromError(ex.Message); }
                }
                case "CONCATENATE":
                {
                    if (fn.Args.Count < 1) return EvaluationResult.FromError("CONCATENATE expects at least 1 argument");
                    try
                    {
                        var parts = new List<string>();
                        foreach (var a in fn.Args) parts.Add(AsText(a));
                        return EvaluationResult.FromText(string.Concat(parts));
                    }
                    catch (Exception ex) { return EvaluationResult.FromError(ex.Message); }
                }
                case "ABS":
                {
                    if (fn.Args.Count != 1) return EvaluationResult.FromError("ABS expects 1 argument");
                    try { var n = GetNumber(fn.Args[0], "ABS"); return EvaluationResult.FromNumber(Math.Abs(n)); }
                    catch (Exception ex) { return EvaluationResult.FromError(ex.Message); }
                }
                case "ROUND":
                {
                    if (fn.Args.Count != 2) return EvaluationResult.FromError("ROUND expects 2 arguments");
                    try
                    {
                        var n = GetNumber(fn.Args[0], "ROUND");
                        int digits = (int)Math.Floor(GetNumber(fn.Args[1], "ROUND"));
                        return EvaluationResult.FromNumber(Math.Round(n, digits, MidpointRounding.AwayFromZero));
                    }
                    catch (Exception ex) { return EvaluationResult.FromError(ex.Message); }
                }
                case "CEILING":
                {
                    if (fn.Args.Count != 1) return EvaluationResult.FromError("CEILING expects 1 argument");
                    try { var n = GetNumber(fn.Args[0], "CEILING"); return EvaluationResult.FromNumber(Math.Ceiling(n)); }
                    catch (Exception ex) { return EvaluationResult.FromError(ex.Message); }
                }
                case "FLOOR":
                {
                    if (fn.Args.Count != 1) return EvaluationResult.FromError("FLOOR expects 1 argument");
                    try { var n = GetNumber(fn.Args[0], "FLOOR"); return EvaluationResult.FromNumber(Math.Floor(n)); }
                    catch (Exception ex) { return EvaluationResult.FromError(ex.Message); }
                }
                case "MOD":
                {
                    if (fn.Args.Count != 2) return EvaluationResult.FromError("MOD expects 2 arguments");
                    try
                    {
                        var a = GetNumber(fn.Args[0], "MOD");
                        var b = GetNumber(fn.Args[1], "MOD");
                        if (b == 0) return EvaluationResult.FromError("DIV/0!");
                        var res = a - b * Math.Floor(a / b);
                        return EvaluationResult.FromNumber(res);
                    }
                    catch (Exception ex) { return EvaluationResult.FromError(ex.Message); }
                }
                case "VLOOKUP":
                {
                    if (fn.Args.Count is not 3 and not 4)
                        return EvaluationResult.FromError("VLOOKUP expects 3 or 4 arguments");
                    var searchVal = EvalNode(fn.Args[0]);
                    if (searchVal.Error != null) return searchVal;
                    if (fn.Args[1] is not RangeNode rng)
                        return EvaluationResult.FromError("VLOOKUP expects a range as 2nd argument");

                    if (!CellAddress.TryParse(rng.A, out int r1, out int c1) || !CellAddress.TryParse(rng.B, out int r2, out int c2))
                        return EvaluationResult.FromError("VLOOKUP bad range");
                    int rStart = Math.Min(r1, r2), rEnd = Math.Max(r1, r2);
                    int cStart = Math.Min(c1, c2), cEnd = Math.Max(c1, c2);
                    int width = cEnd - cStart + 1;
                    var colIndexVal = EvalNode(fn.Args[2]);
                    if (colIndexVal.Error != null) return colIndexVal;
                    if (colIndexVal.Number is not double cold)
                        return EvaluationResult.FromError("VLOOKUP col_index must be a number");
                    int colIndex = (int)Math.Floor(cold);
                    if (colIndex < 1 || colIndex > width)
                        return EvaluationResult.FromError("VLOOKUP col_index out of range");
                    bool exact = true;
                    if (fn.Args.Count == 4)
                    {
                        try { exact = Truthy(fn.Args[3]); }
                        catch (Exception ex) { return EvaluationResult.FromError(ex.Message); }
                    }

                    // Exact match search
                    if (exact)
                    {
                        for (int rr = rStart; rr <= rEnd; rr++)
                        {
                            var firstAddr = CellAddress.ToAddress(rr, cStart);
                            var v = _cellResolver(firstAddr);
                            if (v.Error != null) continue;
                            bool match;
                            if (searchVal.Number is double sn && v.Number is double vn) match = sn == vn;
                            else match = string.Equals(searchVal.ToDisplay(), v.ToDisplay(), StringComparison.OrdinalIgnoreCase);
                            if (match)
                            {
                                var retAddr = CellAddress.ToAddress(rr, cStart + (colIndex - 1));
                                return _cellResolver(retAddr);
                            }
                        }
                        return EvaluationResult.FromError("VLOOKUP not found");
                    }
                    else
                    {
                        // Approximate match: numeric only, last <= search
                        if (searchVal.Number is not double target)
                            return EvaluationResult.FromError("VLOOKUP approximate requires numeric search");
                        int bestRow = -1;
                        double bestVal = double.NegativeInfinity;
                        for (int rr = rStart; rr <= rEnd; rr++)
                        {
                            var firstAddr = CellAddress.ToAddress(rr, cStart);
                            var v = _cellResolver(firstAddr);
                            if (v.Error != null || v.Number is not double vn) continue;
                            if (vn <= target && vn >= bestVal)
                            {
                                bestVal = vn; bestRow = rr;
                            }
                        }
                        if (bestRow == -1) return EvaluationResult.FromError("VLOOKUP not found");
                        var retAddr = CellAddress.ToAddress(bestRow, cStart + (colIndex - 1));
                        return _cellResolver(retAddr);
                    }
                }
                default:
                    return EvaluationResult.FromError($"Unknown function {fn.Name}");
            }
        }

        private static IEnumerable<string> EnumerateRange(string a, string b)
        {
            if (!CellAddress.TryParse(a, out int r1, out int c1) || !CellAddress.TryParse(b, out int r2, out int c2))
                yield break;
            int rStart = Math.Min(r1, r2), rEnd = Math.Max(r1, r2);
            int cStart = Math.Min(c1, c2), cEnd = Math.Max(c1, c2);
            for (int r = rStart; r <= rEnd; r++)
                for (int c = cStart; c <= cEnd; c++)
                    yield return CellAddress.ToAddress(r, c);
        }

        // -------------- Parser -----------------
        private enum TokenType { Number, Ident, Cell, String, Plus, Minus, Star, Slash, Caret, Ampersand, LParen, RParen, Comma, Colon, Eq, NotEq, Lt, Gt, LtEq, GtEq, End }

        private record Token(TokenType Type, string Text, double Number = 0);

        private abstract record Node;
        private record NumberNode(double Value) : Node;
        private record StringNode(string Value) : Node;
        private record CellRefNode(string Address) : Node;
        private record RangeNode(string A, string B) : Node;
        private record UnaryNode(TokenType Op, Node Inner) : Node;
        private record BinaryNode(TokenType Op, Node Left, Node Right) : Node;
        private record FuncCallNode(string Name, List<Node> Args) : Node;

        private class Parser
        {
            private readonly string _s;
            private int _pos;
            private Token _look;

            public Parser(string s)
            {
                _s = s ?? string.Empty;
                _pos = 0;
                _look = NextToken();
            }

            public bool AtEnd => _look.Type == TokenType.End;

            public Node ParseExpression() => ParseConcat();

            private Node ParseConcat()
            {
                var node = ParseComparison();
                while (_look.Type is TokenType.Ampersand)
                {
                    var op = _look.Type; Consume();
                    var rhs = ParseComparison();
                    node = new BinaryNode(op, node, rhs);
                }
                return node;
            }

            private Node ParseComparison()
            {
                var node = ParseAddSub();
                while (_look.Type is TokenType.Eq or TokenType.NotEq or TokenType.Lt or TokenType.Gt or TokenType.LtEq or TokenType.GtEq)
                {
                    var op = _look.Type; Consume();
                    var rhs = ParseAddSub();
                    node = new BinaryNode(op, node, rhs);
                }
                return node;
            }

            private Node ParseAddSub()
            {
                var node = ParseMulDiv();
                while (_look.Type is TokenType.Plus or TokenType.Minus)
                {
                    var op = _look.Type; Consume();
                    var rhs = ParseMulDiv();
                    node = new BinaryNode(op, node, rhs);
                }
                return node;
            }

            private Node ParseMulDiv()
            {
                var node = ParsePower();
                while (_look.Type is TokenType.Star or TokenType.Slash)
                {
                    var op = _look.Type; Consume();
                    var rhs = ParsePower();
                    node = new BinaryNode(op, node, rhs);
                }
                return node;
            }

            private Node ParsePower()
            {
                var node = ParseUnary();
                if (_look.Type == TokenType.Caret)
                {
                    Consume();
                    var rhs = ParsePower(); // right-assoc
                    node = new BinaryNode(TokenType.Caret, node, rhs);
                }
                return node;
            }

            private Node ParseUnary()
            {
                if (_look.Type == TokenType.Plus) { Consume(); return new UnaryNode(TokenType.Plus, ParseUnary()); }
                if (_look.Type == TokenType.Minus) { Consume(); return new UnaryNode(TokenType.Minus, ParseUnary()); }
                return ParsePrimary();
            }

            private Node ParsePrimary()
            {
                if (_look.Type == TokenType.Number) { var v = _look.Number; Consume(); return new NumberNode(v); }
                if (_look.Type == TokenType.String) { var s = _look.Text; Consume(); return new StringNode(s); }
                if (_look.Type == TokenType.Cell)
                {
                    string a = _look.Text; Consume();
                    if (_look.Type == TokenType.Colon)
                    {
                        Consume();
                        if (_look.Type != TokenType.Cell) throw new Exception("Range needs cell after colon");
                        string b = _look.Text; Consume();
                        return new RangeNode(a, b);
                    }
                    return new CellRefNode(a);
                }
                if (_look.Type == TokenType.Ident)
                {
                    string name = _look.Text; Consume();
                    if (_look.Type == TokenType.LParen)
                    {
                        Consume();
                        var args = new List<Node>();
                        if (_look.Type != TokenType.RParen)
                        {
                            while (true)
                            {
                                args.Add(ParseExpression());
                                if (_look.Type == TokenType.Comma) { Consume(); continue; }
                                break;
                            }
                        }
                        if (_look.Type != TokenType.RParen) throw new Exception("Missing )");
                        Consume();
                        return new FuncCallNode(name, args);
                    }
                    throw new Exception("Identifier must be a function");
                }
                if (_look.Type == TokenType.LParen)
                {
                    Consume();
                    var inner = ParseExpression();
                    if (_look.Type != TokenType.RParen) throw new Exception("Missing )");
                    Consume();
                    return inner;
                }
                throw new Exception("Unexpected token");
            }

            private void Consume()
            {
                _look = NextToken();
            }

            private Token NextToken()
            {
                SkipWs();
                if (_pos >= _s.Length) return new Token(TokenType.End, string.Empty);
                char ch = _s[_pos];
                if (char.IsDigit(ch) || (ch == '.' && _pos + 1 < _s.Length && char.IsDigit(_s[_pos + 1])))
                {
                    int start = _pos;
                    _pos++;
                    while (_pos < _s.Length && (char.IsDigit(_s[_pos]) || _s[_pos] == '.')) _pos++;
                    var text = _s.Substring(start, _pos - start);
                    if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double d))
                        throw new Exception("Bad number");
                    return new Token(TokenType.Number, text, d);
                }
                if (char.IsLetter(ch) || ch == '$')
                {
                    int start = _pos;
                    // Optional $ before letters
                    if (_s[_pos] == '$') _pos++;
                    int lettersStart = _pos;
                    if (_pos < _s.Length && char.IsLetter(_s[_pos]))
                    {
                        _pos++;
                        while (_pos < _s.Length && char.IsLetter(_s[_pos])) _pos++;
                    }
                    else
                    {
                        // Not a valid identifier/cell after $
                        throw new Exception($"Unexpected '{_s[_pos-1]}'");
                    }
                    int lettersEnd = _pos;
                    // Optional $ before digits
                    if (_pos < _s.Length && _s[_pos] == '$') _pos++;
                    int digitsStart = _pos;
                    while (_pos < _s.Length && char.IsDigit(_s[_pos])) _pos++;
                    int end = _pos;

                    string letters = _s.Substring(lettersStart, lettersEnd - lettersStart);
                    string digits = _s.Substring(digitsStart, end - digitsStart);
                    if (digits.Length > 0)
                    {
                        string addr = (letters + digits).ToUpperInvariant();
                        if (!CellAddress.TryParse(addr, out _, out _)) throw new Exception("Bad cell ref");
                        return new Token(TokenType.Cell, addr);
                    }
                    else
                    {
                        // No digits: this is an identifier (function name)
                        string ident = _s.Substring(start, lettersEnd - start).Replace("$", string.Empty);
                        return new Token(TokenType.Ident, ident);
                    }
                }

                // String literal
                if (ch == '"')
                {
                    _pos++; // consume opening quote
                    var start = _pos;
                    var sb = new System.Text.StringBuilder();
                    while (_pos < _s.Length)
                    {
                        char ch2 = _s[_pos];
                        if (ch2 == '"')
                        {
                            // check for escaped quote
                            if (_pos + 1 < _s.Length && _s[_pos + 1] == '"')
                            {
                                sb.Append('"');
                                _pos += 2;
                                continue;
                            }
                            _pos++; // consume closing quote
                            return new Token(TokenType.String, sb.ToString());
                        }
                        sb.Append(ch2);
                        _pos++;
                    }
                    throw new Exception("Unterminated string");
                }

                // Multi-char operator lookahead for comparisons
                if (ch == '<')
                {
                    if (_pos + 1 < _s.Length)
                    {
                        char n = _s[_pos + 1];
                        if (n == '=') { _pos += 2; return new Token(TokenType.LtEq, "<="); }
                        if (n == '>') { _pos += 2; return new Token(TokenType.NotEq, "<>"); }
                    }
                    _pos++;
                    return new Token(TokenType.Lt, "<");
                }
                if (ch == '>')
                {
                    if (_pos + 1 < _s.Length && _s[_pos + 1] == '=') { _pos += 2; return new Token(TokenType.GtEq, ">="); }
                    _pos++;
                    return new Token(TokenType.Gt, ">");
                }
                if (ch == '=') { _pos++; return new Token(TokenType.Eq, "="); }

                _pos++;
                return ch switch
                {
                    '+' => new Token(TokenType.Plus, "+"),
                    '-' => new Token(TokenType.Minus, "-"),
                    '*' => new Token(TokenType.Star, "*"),
                    '/' => new Token(TokenType.Slash, "/"),
                    '^' => new Token(TokenType.Caret, "^"),
                    '&' => new Token(TokenType.Ampersand, "&"),
                    '(' => new Token(TokenType.LParen, "("),
                    ')' => new Token(TokenType.RParen, ")"),
                    ',' => new Token(TokenType.Comma, ","),
                    ':' => new Token(TokenType.Colon, ":"),
                    _ => throw new Exception($"Unexpected '{ch}'")
                };
            }

            private void SkipWs()
            {
                while (_pos < _s.Length && char.IsWhiteSpace(_s[_pos])) _pos++;
            }
        }
    }
}
