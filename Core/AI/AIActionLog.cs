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
        }

        private readonly List<Entry> _entries = new();

        public IReadOnlyList<Entry> Entries => _entries;

        public void Record(string prompt, AIPlan plan)
        {
            int cells = 0;
            cells += plan.Commands.OfType<SetValuesCommand>().Sum(c => c.Values.Sum(r => r?.Length ?? 0));
            cells += plan.Commands.OfType<SetFormulaCommand>().Sum(c => c.Formulas.Sum(r => r?.Length ?? 0));
            var summary = string.Join("; ", plan.Commands.Select(c => c.Summarize()));
            _entries.Add(new Entry
            {
                Timestamp = DateTime.Now,
                Prompt = prompt.Length > 120 ? prompt.Substring(0, 120) + "..." : prompt,
                CommandCount = plan.Commands.Count,
                CellCount = cells,
                Summary = summary.Length > 200 ? summary.Substring(0, 200) + "..." : summary
            });
        }

        public void Clear() => _entries.Clear();
    }
}
