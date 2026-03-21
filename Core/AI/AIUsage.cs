using System;

namespace SpreadsheetApp.Core.AI
{
    public sealed class AIUsage
    {
        public int? InputTokens { get; set; }
        public int? OutputTokens { get; set; }
        public int? TotalTokens { get; set; }
        public int? ContextLimit { get; set; }
        public int? RemainingContext => (InputTokens.HasValue && ContextLimit.HasValue) ? Math.Max(0, ContextLimit.Value - InputTokens.Value) : (int?)null;
    }
}

