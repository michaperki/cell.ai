using System;
using System.Collections.Generic;
using System.Linq;

namespace SpreadsheetApp.Core.AI
{
    public sealed class AIActionLog
    {
        public sealed class Entry
        {
            public DateTime Timestamp { get; set; }
            public string Prompt { get; set; } = string.Empty;
            public int CommandCount { get; set; }
            public int CellCount { get; set; }
            public string Summary { get; set; } = string.Empty;
            public string? Provider { get; set; }
            public string? Model { get; set; }
            public int? InputTokens { get; set; }
            public int? OutputTokens { get; set; }
            public int? TotalTokens { get; set; }
            public int? LatencyMs { get; set; }
        }

        private readonly List<Entry> _entries = new();

        public IReadOnlyList<Entry> Entries => _entries;

        public void Record(string prompt, AIPlan plan)
        {
            string p = prompt ?? string.Empty;
            var pl = plan ?? new AIPlan();
            var commands = pl.Commands ?? new System.Collections.Generic.List<IAICommand>();
            int cells = 0;
            try { cells += commands.OfType<SetValuesCommand>().Sum(c => c.Values.Sum(r => r?.Length ?? 0)); } catch { }
            try { cells += commands.OfType<SetFormulaCommand>().Sum(c => c.Formulas.Sum(r => r?.Length ?? 0)); } catch { }
            var summary = string.Join("; ", commands.Select(c => c.Summarize()));
            _entries.Add(new Entry
            {
                Timestamp = DateTime.Now,
                Prompt = p.Length > 120 ? p.Substring(0, 120) + "..." : p,
                CommandCount = commands.Count,
                CellCount = cells,
                Summary = summary.Length > 200 ? summary.Substring(0, 200) + "..." : summary,
                Provider = pl.Provider,
                Model = pl.Model,
                InputTokens = pl.Usage?.InputTokens,
                OutputTokens = pl.Usage?.OutputTokens,
                TotalTokens = pl.Usage?.TotalTokens,
                LatencyMs = pl.LatencyMs
            });
        }

        public void Clear() => _entries.Clear();
    }
}
