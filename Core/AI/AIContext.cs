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
}
