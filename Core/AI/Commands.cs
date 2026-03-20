using System.Collections.Generic;

namespace SpreadsheetApp.Core.AI
{
    public enum AICommandType { SetValues, SetTitle, CreateSheet, ClearRange, RenameSheet, SetFormula, SortRange, InsertRows, DeleteRows, InsertCols, DeleteCols, DeleteSheet, CopyRange, MoveRange, SetFormat, SetValidation, SetConditionalFormat, TransformRange }

    public interface IAICommand
    {
        AICommandType Type { get; }
        string Summarize();
    }

    public sealed class SetValuesCommand : IAICommand
    {
        public AICommandType Type => AICommandType.SetValues;
        public int StartRow { get; set; }
        public int StartCol { get; set; }
        public string[][] Values { get; set; } = new string[0][];
        public string? Rationale { get; set; }
        public string Summarize() => $"Set {Values.Length}x{(Values.Length>0?Values[0].Length:0)} values at {StartRow+1},{StartCol+1}";
    }

    public sealed class SetFormulaCommand : IAICommand
    {
        public AICommandType Type => AICommandType.SetFormula;
        public int StartRow { get; set; }
        public int StartCol { get; set; }
        public string[][] Formulas { get; set; } = new string[0][];
        public string? Rationale { get; set; }
        public string Summarize() => $"Set {Formulas.Length}x{(Formulas.Length>0?Formulas[0].Length:0)} formulas at {StartRow+1},{StartCol+1}";
    }

    public sealed class SetTitleCommand : IAICommand
    {
        public AICommandType Type => AICommandType.SetTitle;
        public int StartRow { get; set; }
        public int StartCol { get; set; }
        public int Rows { get; set; } = 1;
        public int Cols { get; set; } = 1;
        public string Text { get; set; } = string.Empty;
        public string? Rationale { get; set; }
        public string Summarize() => $"Title '{Text}' at {StartRow+1},{StartCol+1} ({Rows}x{Cols})";
    }

    public sealed class AIPlan
    {
        public List<IAICommand> Commands { get; } = new();
        // Optional raw JSON from provider for debugging
        public string? RawJson { get; set; }
        // Optional constructed user prompt string for debugging
        public string? RawUser { get; set; }
        // Optional system prompt (schema/rules) for debugging
        public string? RawSystem { get; set; }
        // Optional: query intents for observation phase (two-phase agent loop)
        public List<IQueryIntent> Queries { get; } = new();
    }

    public enum QueryIntentType { SelectionSummary, ProfileColumn, UniqueValues, SampleRows, DescribeColumn, CountWhere }

    public interface IQueryIntent
    {
        QueryIntentType Type { get; }
        string Summarize();
    }

    public sealed class SelectionSummaryQuery : IQueryIntent
    {
        public QueryIntentType Type => QueryIntentType.SelectionSummary;
        public string Summarize() => "selection_summary";
    }

    public sealed class ProfileColumnQuery : IQueryIntent
    {
        public QueryIntentType Type => QueryIntentType.ProfileColumn;
        public int ColumnIndex { get; set; } // 0-based absolute column index
        public int? Rows { get; set; }
        public string Summarize() => $"profile_column {SpreadsheetApp.Core.CellAddress.ColumnIndexToName(ColumnIndex)}";
    }

    public sealed class DescribeColumnQuery : IQueryIntent
    {
        public QueryIntentType Type => QueryIntentType.DescribeColumn;
        public int ColumnIndex { get; set; } // 0-based absolute column index
        public int? Rows { get; set; }
        public string Summarize() => $"describe_column {SpreadsheetApp.Core.CellAddress.ColumnIndexToName(ColumnIndex)}";
    }

    public sealed class CountWhereQuery : IQueryIntent
    {
        public sealed class Filter
        {
            public int ColumnIndex { get; set; } // 0-based absolute column index
            public string Op { get; set; } = "eq"; // eq|ne|gt|ge|lt|le|contains|not_contains
            public string Value { get; set; } = string.Empty;
        }

        public QueryIntentType Type => QueryIntentType.CountWhere;
        public System.Collections.Generic.List<Filter> Filters { get; set; } = new();
        public string Summarize()
        {
            var parts = new System.Collections.Generic.List<string>();
            foreach (var f in Filters)
            {
                parts.Add($"{SpreadsheetApp.Core.CellAddress.ColumnIndexToName(f.ColumnIndex)} {f.Op} '{f.Value}'");
            }
            return $"count_where {string.Join(" & ", parts)}";
        }
    }

    public sealed class UniqueValuesQuery : IQueryIntent
    {
        public QueryIntentType Type => QueryIntentType.UniqueValues;
        public int ColumnIndex { get; set; } // 0-based absolute column index
        public int TopK { get; set; } = 10;
        public string Summarize() => $"unique_values {SpreadsheetApp.Core.CellAddress.ColumnIndexToName(ColumnIndex)} top={TopK}";
    }

    public sealed class SampleRowsQuery : IQueryIntent
    {
        public QueryIntentType Type => QueryIntentType.SampleRows;
        public int Rows { get; set; } = 5;
        public int Cols { get; set; } = 6;
        public string Summarize() => $"sample_rows {Rows}x{Cols}";
    }

    public sealed class CreateSheetCommand : IAICommand
    {
        public AICommandType Type => AICommandType.CreateSheet;
        public string Name { get; set; } = "New Sheet";
        public string? Rationale { get; set; }
        public string Summarize() => $"Create sheet '{Name}'";
    }

    public sealed class ClearRangeCommand : IAICommand
    {
        public AICommandType Type => AICommandType.ClearRange;
        public int StartRow { get; set; }
        public int StartCol { get; set; }
        public int Rows { get; set; } = 1;
        public int Cols { get; set; } = 1;
        public string? Rationale { get; set; }
        public string Summarize() => $"Clear {Rows}x{Cols} at {StartRow+1},{StartCol+1}";
    }

    public sealed class RenameSheetCommand : IAICommand
    {
        public AICommandType Type => AICommandType.RenameSheet;
        // One of Index (1-based) or OldName can be provided; if neither, applies to active.
        public int? Index1 { get; set; }
        public string? OldName { get; set; }
        public string NewName { get; set; } = "Sheet";
        public string? Rationale { get; set; }
        public string Summarize() => $"Rename sheet {(Index1.HasValue ? Index1.ToString() : (OldName ?? "(active)"))} to '{NewName}'";
    }

    public sealed class SortRangeCommand : IAICommand
    {
        public AICommandType Type => AICommandType.SortRange;
        public int StartRow { get; set; }
        public int StartCol { get; set; }
        public int Rows { get; set; } = 1;
        public int Cols { get; set; } = 1;
        // 0-based index of column within the sheet (not relative). Filled by planner from letter.
        public int SortCol { get; set; }
        public string Order { get; set; } = "asc"; // asc | desc
        public bool HasHeader { get; set; } = false;
        public string? Rationale { get; set; }
        public string Summarize() => $"Sort {Rows}x{Cols} at {StartRow+1},{StartCol+1} by {SpreadsheetApp.Core.CellAddress.ColumnIndexToName(SortCol)} {Order}{(HasHeader?", header":"")}";
    }

    public sealed class InsertRowsCommand : IAICommand
    {
        public AICommandType Type => AICommandType.InsertRows;
        public int At { get; set; }     // 0-based row index to insert before
        public int Count { get; set; } = 1;
        public string? Rationale { get; set; }
        public string Summarize() => $"Insert {Count} row(s) at row {At + 1}";
    }

    public sealed class DeleteRowsCommand : IAICommand
    {
        public AICommandType Type => AICommandType.DeleteRows;
        public int At { get; set; }     // 0-based row index to start deleting
        public int Count { get; set; } = 1;
        public string? Rationale { get; set; }
        public string Summarize() => $"Delete {Count} row(s) at row {At + 1}";
    }

    public sealed class InsertColsCommand : IAICommand
    {
        public AICommandType Type => AICommandType.InsertCols;
        public int At { get; set; }     // 0-based column index to insert before
        public int Count { get; set; } = 1;
        public string? Rationale { get; set; }
        public string Summarize() => $"Insert {Count} column(s) at col {SpreadsheetApp.Core.CellAddress.ColumnIndexToName(At)}";
    }

    public sealed class DeleteColsCommand : IAICommand
    {
        public AICommandType Type => AICommandType.DeleteCols;
        public int At { get; set; }     // 0-based column index to start deleting
        public int Count { get; set; } = 1;
        public string? Rationale { get; set; }
        public string Summarize() => $"Delete {Count} column(s) at col {SpreadsheetApp.Core.CellAddress.ColumnIndexToName(At)}";
    }

    public sealed class DeleteSheetCommand : IAICommand
    {
        public AICommandType Type => AICommandType.DeleteSheet;
        // One of Index (1-based) or Name can be provided; if neither, applies to active.
        public int? Index1 { get; set; }
        public string? Name { get; set; }
        public string? Rationale { get; set; }
        public string Summarize() => $"Delete sheet {(Index1.HasValue ? Index1.ToString() : (Name ?? "(active)"))}";
    }

    public sealed class CopyRangeCommand : IAICommand
    {
        public AICommandType Type => AICommandType.CopyRange;
        public int StartRow { get; set; }
        public int StartCol { get; set; }
        public int Rows { get; set; } = 1;
        public int Cols { get; set; } = 1;
        public int DestRow { get; set; }
        public int DestCol { get; set; }
        public string? Rationale { get; set; }
        public string Summarize() => $"Copy {Rows}x{Cols} from {StartRow+1},{StartCol+1} to {DestRow+1},{DestCol+1}";
    }

    public sealed class MoveRangeCommand : IAICommand
    {
        public AICommandType Type => AICommandType.MoveRange;
        public int StartRow { get; set; }
        public int StartCol { get; set; }
        public int Rows { get; set; } = 1;
        public int Cols { get; set; } = 1;
        public int DestRow { get; set; }
        public int DestCol { get; set; }
        public string? Rationale { get; set; }
        public string Summarize() => $"Move {Rows}x{Cols} from {StartRow+1},{StartCol+1} to {DestRow+1},{DestCol+1}";
    }

    public sealed class SetFormatCommand : IAICommand
    {
        public AICommandType Type => AICommandType.SetFormat;
        public int StartRow { get; set; }
        public int StartCol { get; set; }
        public int Rows { get; set; } = 1;
        public int Cols { get; set; } = 1;
        public bool? Bold { get; set; }
        public int? ForeColorArgb { get; set; } // ARGB int, e.g., 0xFFRRGGBB
        public int? BackColorArgb { get; set; }
        public string? NumberFormat { get; set; }
        public string? HAlign { get; set; } // left|center|right
        public string? Rationale { get; set; }
        public string Summarize()
            => $"Format {Rows}x{Cols} at {StartRow+1},{StartCol+1} (bold={(Bold?.ToString() ?? "-")}, fore={(ForeColorArgb?.ToString("X") ?? "-")}, back={(BackColorArgb?.ToString("X") ?? "-")}, num={(NumberFormat ?? "-")}, align={(HAlign ?? "-")})";
    }

    public sealed class SetValidationCommand : IAICommand
    {
        public AICommandType Type => AICommandType.SetValidation;
        public int StartRow { get; set; }
        public int StartCol { get; set; }
        public int Rows { get; set; } = 1;
        public int Cols { get; set; } = 1;
        public string Mode { get; set; } = "none"; // none|list|number_between
        public bool AllowEmpty { get; set; } = true;
        public double? Min { get; set; }
        public double? Max { get; set; }
        public string[]? AllowedList { get; set; }
        public string? Rationale { get; set; }
        public string Summarize() => $"Validation {Mode} at {StartRow+1},{StartCol+1} ({Rows}x{Cols})";
    }

    public sealed class SetConditionalFormatCommand : IAICommand
    {
        public AICommandType Type => AICommandType.SetConditionalFormat;
        public int StartRow { get; set; }
        public int StartCol { get; set; }
        public int Rows { get; set; } = 1;
        public int Cols { get; set; } = 1;
        public string Operator { get; set; } = ">"; // >,>=,<,<=,==,!=
        public double Threshold { get; set; }
        public bool? Bold { get; set; }
        public int? ForeColorArgb { get; set; }
        public int? BackColorArgb { get; set; }
        public string? NumberFormat { get; set; }
        public string? HAlign { get; set; }
        public string? Rationale { get; set; }
        public string Summarize() => $"CondFmt {Operator}{Threshold} at {StartRow+1},{StartCol+1} ({Rows}x{Cols})";
    }

    public sealed class TransformRangeCommand : IAICommand
    {
        public AICommandType Type => AICommandType.TransformRange;
        public int StartRow { get; set; }
        public int StartCol { get; set; }
        public int Rows { get; set; } = 1;
        public int Cols { get; set; } = 1;
        public string Op { get; set; } = "trim"; // trim|proper|upper|lower|strip_punct|normalize_city
        public string? Rationale { get; set; }
        public string Summarize() => $"Transform '{Op}' at {StartRow+1},{StartCol+1} ({Rows}x{Cols})";
    }
}
