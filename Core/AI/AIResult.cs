namespace SpreadsheetApp.Core.AI
{
    public sealed class AIResult
    {
        // Two-dimensional cells [row][col]
        public string[][] Cells { get; set; } = new string[0][];
        // Provider/model metadata and token usage (if available)
        public string? Provider { get; set; }
        public string? Model { get; set; }
        public AIUsage? Usage { get; set; }
        public int? LatencyMs { get; set; }
    }
}
