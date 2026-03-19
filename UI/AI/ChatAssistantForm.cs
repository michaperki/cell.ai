using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using SpreadsheetApp.Core.AI;

namespace SpreadsheetApp.UI.AI
{
    public sealed class ChatAssistantForm : Form
    {
        private readonly IChatPlanner _planner;
        private readonly Func<AIContext> _getContext;
        private readonly Action<AIPlan> _applyPlan;
        private readonly TextBox _input = new() { Dock = DockStyle.Top, Multiline = true, Height = 60 };
        private readonly Button _btnPlan = new() { Text = "Plan", Dock = DockStyle.Top, Height = 28 };
        private readonly ListBox _lst = new() { Dock = DockStyle.Fill };
        private readonly Button _btnApply = new() { Text = "Apply", Dock = DockStyle.Bottom, Height = 32, Enabled = false };
        private readonly Button _btnClose = new() { Text = "Close", Dock = DockStyle.Bottom, Height = 28 };
        private AIPlan? _currentPlan;

        public ChatAssistantForm(IChatPlanner planner, Func<AIContext> getContext, Action<AIPlan> applyPlan)
        {
            _planner = planner;
            _getContext = getContext;
            _applyPlan = applyPlan;
            Text = "AI Chat Assistant";
            StartPosition = FormStartPosition.CenterParent;
            Width = 420; Height = 480;
            var panel = new Panel { Dock = DockStyle.Fill };
            panel.Controls.Add(_lst);
            panel.Controls.Add(_btnApply);
            panel.Controls.Add(_btnClose);
            panel.Controls.Add(_btnPlan);
            panel.Controls.Add(_input);
            Controls.Add(panel);

            _btnPlan.Click += async (_, __) => await DoPlanAsync();
            _btnApply.Click += (_, __) => { if (_currentPlan != null) { _applyPlan(_currentPlan); Close(); } };
            _btnClose.Click += (_, __) => Close();
            AcceptButton = _btnPlan;
        }

        private async Task DoPlanAsync()
        {
            _btnPlan.Enabled = false; _btnApply.Enabled = false; _lst.Items.Clear(); _currentPlan = null;
            try
            {
                using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(10));
                var ctx = _getContext();
                var plan = await _planner.PlanAsync(ctx, _input.Text ?? string.Empty, cts.Token).ConfigureAwait(true);
                _currentPlan = plan;
                if (plan.Commands.Count == 0)
                {
                    _lst.Items.Add("No changes suggested.");
                }
                else
                {
                    foreach (var cmd in plan.Commands) _lst.Items.Add(cmd.Summarize());
                    int writeCount = plan.Commands.OfType<SetValuesCommand>().Sum(c => c.Values.Length * (c.Values.Length > 0 ? c.Values[0].Length : 0));
                    _lst.Items.Add($"Total writes: {writeCount}");
                    _btnApply.Enabled = true;
                }
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
            }
        }
    }
}

