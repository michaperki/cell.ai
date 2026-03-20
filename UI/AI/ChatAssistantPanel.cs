using System;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using SpreadsheetApp.Core.AI;

namespace SpreadsheetApp.UI.AI
{
    // Dockable chat assistant panel (reuses ChatAssistantForm logic in a UserControl)
    public sealed class ChatAssistantPanel : UserControl
    {
        private readonly IChatPlanner _planner;
        private readonly Func<AIContext> _getContext;
        private readonly Action<AIPlan> _applyPlan;

        private readonly TextBox _input = new() { Dock = DockStyle.Top, Multiline = true, Height = 60 };
        private readonly Button _btnPlan = new() { Text = "Plan", Dock = DockStyle.Top, Height = 28 };
        private readonly Button _btnRevise = new() { Text = "Revise", Dock = DockStyle.Top, Height = 24 };
        private readonly Button _btnReset = new() { Text = "Reset History", Dock = DockStyle.Top, Height = 24 };
        private readonly Label _lblStatus = new() { Dock = DockStyle.Top, Height = 18, Text = string.Empty, TextAlign = ContentAlignment.MiddleLeft, ForeColor = Color.DimGray, Visible = false };
        private readonly ListBox _lst = new() { Dock = DockStyle.Fill };
        private readonly Button _btnApply = new() { Text = "Apply", Dock = DockStyle.Bottom, Height = 28, Enabled = false };

        private AIPlan? _currentPlan;
        private readonly System.Collections.Generic.List<ChatMessage> _history = new();

        public ChatAssistantPanel(IChatPlanner planner, Func<AIContext> getContext, Action<AIPlan> applyPlan, string? initialPrompt = null, bool autoPlan = false)
        {
            _planner = planner;
            _getContext = getContext;
            _applyPlan = applyPlan;

            Dock = DockStyle.Fill;
            Padding = new Padding(6);
            BackColor = Color.White;

            var container = new Panel { Dock = DockStyle.Fill };
            container.Controls.Add(_lst);
            container.Controls.Add(_btnApply);
            container.Controls.Add(_btnPlan);
            container.Controls.Add(_btnRevise);
            container.Controls.Add(_lblStatus);
            container.Controls.Add(_btnReset);
            container.Controls.Add(_input);
            Controls.Add(container);

            _btnPlan.Click += async (_, __) => await DoPlanAsync();
            _btnRevise.Click += async (_, __) =>
            {
                var feedback = _input.Text ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(feedback))
                {
                    _history.Add(new ChatMessage { Role = "user", Content = feedback });
                    if (_history.Count > 10) _history.RemoveRange(0, _history.Count - 10);
                }
                await DoPlanAsync();
            };
            _btnReset.Click += (_, __) => { _history.Clear(); _lst.Items.Add("History cleared."); };
            _btnApply.Click += (_, __) =>
            {
                if (_currentPlan != null)
                {
                    try
                    {
                        _applyPlan(_currentPlan);
                        var summary = string.Join("; ", _currentPlan.Commands.Select(c => c.Summarize()));
                        _lst.Items.Add($"Applied: {summary}");
                    }
                    catch (Exception ex)
                    {
                        _lst.Items.Add($"Apply error: {ex.Message}");
                    }
                    finally
                    {
                        _currentPlan = null;
                        _btnApply.Enabled = false;
                        _input.Clear();
                        _input.Focus();
                    }
                }
            };

            if (!string.IsNullOrWhiteSpace(initialPrompt)) _input.Text = initialPrompt;
            if (autoPlan) _ = DoPlanAsync();
        }

        public void FocusInput()
        {
            try { _input.Focus(); } catch { }
        }

        public void SetPrompt(string text, bool autoPlan)
        {
            _input.Text = text ?? string.Empty;
            if (autoPlan) _ = DoPlanAsync();
        }

        private async Task DoPlanAsync()
        {
            _btnPlan.Enabled = false; _btnApply.Enabled = false; _currentPlan = null;
            _lblStatus.Text = "Thinking..."; _lblStatus.Visible = true;
            int planningMarkerIndex = _lst.Items.Add("Planning...");
            try
            {
                int timeoutSec = 30;
                try { var s = Environment.GetEnvironmentVariable("AI_PLAN_TIMEOUT_SEC"); if (!string.IsNullOrWhiteSpace(s)) timeoutSec = Math.Max(5, int.Parse(s)); } catch { }
                using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(timeoutSec));
                var ctx = _getContext();
                // Include rolling conversation
                ctx.Conversation = new System.Collections.Generic.List<ChatMessage>(_history);
                var plan = await _planner.PlanAsync(ctx, _input.Text ?? string.Empty, cts.Token).ConfigureAwait(true);
                _currentPlan = plan;
                _lst.Items.Clear();
                if (plan.Commands.Count == 0)
                {
                    _lst.Items.Add("No changes suggested.");
                }
                else
                {
                    foreach (var cmd in plan.Commands) _lst.Items.Add(cmd.Summarize());
                    int writeCount = 0;
                    writeCount += plan.Commands.OfType<SetValuesCommand>().Sum(c => c.Values.Length * (c.Values.Length > 0 ? c.Values[0].Length : 0));
                    writeCount += plan.Commands.OfType<SetFormulaCommand>().Sum(c => c.Formulas.Length * (c.Formulas.Length > 0 ? c.Formulas[0].Length : 0));
                    _lst.Items.Add($"Total writes: {writeCount}");
                    _btnApply.Enabled = true;
                }
                // Update conversation history (keep last 10 entries)
                var userMsg = new ChatMessage { Role = "user", Content = _input.Text ?? string.Empty };
                var asstSummary = string.Join("; ", plan.Commands.Select(c => c.Summarize()));
                var asstMsg = new ChatMessage { Role = "assistant", Content = asstSummary };
                _history.Add(userMsg);
                _history.Add(asstMsg);
                if (_history.Count > 10) _history.RemoveRange(0, _history.Count - 10);
            }
            catch (OperationCanceledException)
            {
                _lst.Items.Add("Planning canceled or timed out.");
            }
            catch (Exception ex)
            {
                _lst.Items.Add($"Error: {ex.Message}");
            }
            finally
            {
                _btnPlan.Enabled = true;
                _lblStatus.Visible = false;
            }
        }
    }
}

