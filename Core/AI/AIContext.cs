namespace SpreadsheetApp.Core.AI
{
    public sealed class AIContext
    {
        public string SheetName { get; set; } = "Sheet1";
        public int StartRow { get; set; }
        public int StartCol { get; set; }
        public int Rows { get; set; } = 1;
        public int Cols { get; set; } = 1;
        public string? Title { get; set; }
        public string Prompt { get; set; } = string.Empty;

        // Optional: selection content (values) and nearby window for planning
        public string[][]? SelectionValues { get; set; }
        public string[][]? NearbyValues { get; set; }

        // Optional: workbook-level summary (sheet names, used sizes, headers)
        public SheetSummary[]? Workbook { get; set; }

        // Optional: rolling conversation for multi-turn chat
        public System.Collections.Generic.List<ChatMessage>? Conversation { get; set; }

        // Optional: explicit command gating for the planner (e.g., ["set_values"]) 
        public string[]? AllowedCommands { get; set; }

        // Optional: write policy for the active selection
        public SelectionWritePolicy? WritePolicy { get; set; }

        // Optional: concise schema for the active selection (headers and types)
        public ColumnSchema[]? Schema { get; set; }

        // Optional: strict selection fencing. When true, the planner/UI should avoid and/or drop any writes outside the selection bounds.
        public bool SelectionHardMode { get; set; }

        // Optional: when true, the planner should produce only observation/query intents instead of write commands.
        public bool RequestQueriesOnly { get; set; }

        // Optional: when true, Ask mode expects a textual answer rather than a write plan.
        public bool AnswerOnly { get; set; }
    }

    public sealed class ChatMessage
    {
        public string Role { get; set; } = "user"; // user | assistant
        public string Content { get; set; } = string.Empty;
    }

    public sealed class SheetSummary
    {
        public string Name { get; set; } = "Sheet";
        public int UsedRows { get; set; }
        public int UsedCols { get; set; }
        public string[]? HeaderRow { get; set; }
        public int HeaderRowIndex { get; set; } = -1;
        public int DataRowCountExcludingHeader { get; set; }
        public string? UsedTopLeft { get; set; }
        public string? UsedBottomRight { get; set; }
    }

    public sealed class SelectionWritePolicy
    {
        // Absolute 0-based column indices allowed for writes for this operation
        public int[]? WritableColumns { get; set; }
        // Input column index if applicable (e.g., first column of the table)
        public int? InputColumnIndex { get; set; }
        // If true, writes to InputColumn are allowed for empty rows within the selection
        public bool AllowInputWritesForEmptyRows { get; set; }
        // If true, writes to InputColumn are allowed for rows that already contain data
        public bool AllowInputWritesForExistingRows { get; set; }
        // Header row is read-only unless explicitly targeted
        public bool HeaderRowReadOnly { get; set; } = true;
    }

    public sealed class ColumnSchema
    {
        public int ColumnIndex { get; set; }
        public string ColumnLetter { get; set; } = "A";
        public string? Name { get; set; }
        public string Type { get; set; } = "text"; // text | number | formula | date
        public bool AllowEmpty { get; set; } = true;
    }
}
