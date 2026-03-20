using System;
using System.Drawing;
using System.Windows.Forms;
using SpreadsheetApp.Core.AI;

namespace SpreadsheetApp.UI.AI
{
    public sealed class ChatAssistantForm : Form
    {
        private readonly ChatAssistantPanel _panel;

        public ChatAssistantForm(IChatPlanner planner, ChatSession session, Func<AIContext> getContext, Action<AIPlan> applyPlan, string? initialPrompt = null, bool autoPlan = false)
        {
            Text = "AI Chat Assistant";
            StartPosition = FormStartPosition.CenterParent;
            Width = 420; Height = 480;

            _panel = new ChatAssistantPanel(planner, session, getContext, applyPlan, initialPrompt, autoPlan);
            _panel.Dock = DockStyle.Fill;

            var btnClose = new Button { Text = "Close", Dock = DockStyle.Bottom, Height = 28 };
            btnClose.Click += (_, __) => Close();

            var container = new Panel { Dock = DockStyle.Fill };
            container.Controls.Add(_panel);
            container.Controls.Add(btnClose);
            Controls.Add(container);
        }
    }
}
