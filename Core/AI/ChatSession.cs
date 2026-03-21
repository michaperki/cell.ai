using System.Collections.Generic;

namespace SpreadsheetApp.Core.AI
{
    // Typed entries for a unified chat thread
    public enum ChatEntryType { User, Answer, ActionProposal, AppliedSummary, Observation }

    public sealed class ChatEntry
    {
        public ChatEntryType Type { get; set; }
        public string Content { get; set; } = string.Empty; // user text, answer, or summary
        public AIPlan? Proposal { get; set; }               // filled for ActionProposal
        public int ProposalVersion { get; set; }            // increments on Revise
        public string? Meta { get; set; }                   // optional small metadata string
    }

    public sealed class ChatSession
    {
        private readonly List<ChatMessage> _history = new();
        public IReadOnlyList<ChatMessage> History => _history;

        private readonly List<ChatEntry> _thread = new();
        public IReadOnlyList<ChatEntry> Thread => _thread;

        public void Clear()
        {
            _history.Clear();
            _thread.Clear();
        }

        // Back-compat helpers (used by planners)
        public void AddUser(string content)
        {
            if (content == null) content = string.Empty;
            _history.Add(new ChatMessage { Role = "user", Content = content });
            _thread.Add(new ChatEntry { Type = ChatEntryType.User, Content = content });
            Trim();
        }

        public void AddAssistant(string content)
        {
            if (content == null) content = string.Empty;
            _history.Add(new ChatMessage { Role = "assistant", Content = content });
            _thread.Add(new ChatEntry { Type = ChatEntryType.Answer, Content = content });
            Trim();
        }

        public void AddAnswer(string text, string? meta = null)
        {
            _history.Add(new ChatMessage { Role = "assistant", Content = text ?? string.Empty });
            _thread.Add(new ChatEntry { Type = ChatEntryType.Answer, Content = text ?? string.Empty, Meta = meta });
            Trim();
        }

        public void AddObservation(string text)
        {
            _thread.Add(new ChatEntry { Type = ChatEntryType.Observation, Content = text ?? string.Empty });
            Trim();
        }

        public int AddProposal(AIPlan plan, int version = 1, string? meta = null)
        {
            _thread.Add(new ChatEntry { Type = ChatEntryType.ActionProposal, Proposal = plan, ProposalVersion = version, Meta = meta, Content = plan != null ? BuildPlanSummary(plan) : string.Empty });
            Trim();
            return _thread.Count - 1; // index
        }

        public void ReplaceProposal(int index, AIPlan newPlan, int newVersion)
        {
            if (index < 0 || index >= _thread.Count) return;
            var e = _thread[index];
            if (e.Type != ChatEntryType.ActionProposal) return;
            e.Proposal = newPlan; e.ProposalVersion = newVersion; e.Content = BuildPlanSummary(newPlan);
        }

        public void AddAppliedSummary(AIPlan plan)
        {
            _thread.Add(new ChatEntry { Type = ChatEntryType.AppliedSummary, Content = BuildPlanSummary(plan) });
            Trim();
        }

        private static string BuildPlanSummary(AIPlan plan)
        {
            try
            {
                if (plan == null || plan.Commands == null || plan.Commands.Count == 0)
                    return string.Empty;
                var parts = new System.Collections.Generic.List<string>();
                foreach (var c in plan.Commands)
                {
                    parts.Add(c.Summarize());
                }
                return string.Join("; ", parts);
            }
            catch { return string.Empty; }
        }

        private void Trim()
        {
            // Keep last ~10 messages and ~30 thread entries to stay light
            const int maxHist = 10;
            const int maxThread = 30;
            if (_history.Count > maxHist)
            {
                _history.RemoveRange(0, _history.Count - maxHist);
            }
            if (_thread.Count > maxThread)
            {
                _thread.RemoveRange(0, _thread.Count - maxThread);
            }
        }
    }
}
