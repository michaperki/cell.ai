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
    }
}

