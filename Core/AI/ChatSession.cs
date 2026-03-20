using System.Collections.Generic;

namespace SpreadsheetApp.Core.AI
{
    public sealed class ChatSession
    {
        private readonly List<ChatMessage> _history = new();
        public IReadOnlyList<ChatMessage> History => _history;

        public void Clear() => _history.Clear();

        public void AddUser(string content)
        {
            if (content == null) content = string.Empty;
            _history.Add(new ChatMessage { Role = "user", Content = content });
            Trim();
        }

        public void AddAssistant(string content)
        {
            if (content == null) content = string.Empty;
            _history.Add(new ChatMessage { Role = "assistant", Content = content });
            Trim();
        }

        private void Trim()
        {
            // Keep last 10 messages (5 exchanges) for light weight context
            const int max = 10;
            if (_history.Count > max)
            {
                _history.RemoveRange(0, _history.Count - max);
            }
        }
    }
}

