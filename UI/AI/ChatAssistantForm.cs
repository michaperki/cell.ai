using System;
using System.Linq;
using System.Drawing;
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
        private readonly TextBox _policy = new() { Dock = DockStyle.Top, Multiline = true, Height = 56, ReadOnly = true, BackColor = SystemColors.Info, Font = new Font("Consolas", 8.5f) };
        private readonly Button _btnPlan = new() { Text = "Plan", Dock = DockStyle.Top, Height = 28 };
        private readonly Button _btnRevise = new() { Text = "Revise", Dock = DockStyle.Top, Height = 24 };
        private readonly Button _btnReset = new() { Text = "Reset History", Dock = DockStyle.Top, Height = 24 };
        private readonly Label _lblStatus = new() { Dock = DockStyle.Top, Height = 18, Text = string.Empty, TextAlign = ContentAlignment.MiddleLeft, ForeColor = Color.DimGray, Visible = false };
        private readonly ListBox _lst = new() { Dock = DockStyle.Fill };
        private readonly Button _btnApply = new() { Text = "Apply", Dock = DockStyle.Bottom, Height = 32, Enabled = false };
        private readonly Button _btnClose = new() { Text = "Close", Dock = DockStyle.Bottom, Height = 28 };
        private AIPlan? _currentPlan;
        private readonly System.Collections.Generic.List<ChatMessage> _history = new();
        private readonly string? _initialPrompt;
        private readonly bool _autoPlan;

        public ChatAssistantForm(IChatPlanner planner, Func<AIContext> getContext, Action<AIPlan> applyPlan, string? initialPrompt = null, bool autoPlan = false)
        {
            _planner = planner;
            _getContext = getContext;
            _applyPlan = applyPlan;
            _initialPrompt = initialPrompt;
            _autoPlan = autoPlan;
            Text = "AI Chat Assistant";
            StartPosition = FormStartPosition.CenterParent;
            Width = 420; Height = 480;
            var panel = new Panel { Dock = DockStyle.Fill };
            panel.Controls.Add(_lst);
            panel.Controls.Add(_btnApply);
            panel.Controls.Add(_btnClose);
            panel.Controls.Add(_btnPlan);
            panel.Controls.Add(_btnRevise);
            panel.Controls.Add(_lblStatus);
            panel.Controls.Add(_btnReset);
            panel.Controls.Add(_policy);
            panel.Controls.Add(_input);
            Controls.Add(panel);

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
            _btnClose.Click += (_, __) => Close();
            AcceptButton = _btnPlan;

            if (!string.IsNullOrWhiteSpace(_initialPrompt)) _input.Text = _initialPrompt;
            Shown += async (_, __) =>
            {
                try { var ctx0 = _getContext(); _policy.Text = BuildPolicyPreview(ctx0); } catch { }
                if (_autoPlan) await DoPlanAsync();
            };
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
                try { _policy.Text = BuildPolicyPreview(ctx); } catch { }
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
                // Update conversation history (keep last 6 messages)
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

        private static string BuildPolicyPreview(AIContext ctx)
        {
            var sb = new System.Text.StringBuilder();
            // Selection summary
            string start = SpreadsheetApp.Core.CellAddress.ToAddress(ctx.StartRow, ctx.StartCol);
            sb.Append($"Selection {start} · {Math.Max(1, ctx.Rows)}x{Math.Max(1, ctx.Cols)}");
            // Allowed commands
            try
            {
                if (ctx.AllowedCommands != null && ctx.AllowedCommands.Length > 0)
                {
                    sb.Append("  |  Allowed: ");
                    for (int i = 0; i < ctx.AllowedCommands.Length; i++)
                    {
                        if (i > 0) sb.Append(',');
                        sb.Append(ctx.AllowedCommands[i]);
                    }
                }
            }
            catch { }
            // Writable columns and input policy
            try
            {
                var wp = ctx.WritePolicy;
                if (wp != null)
                {
                    if (wp.WritableColumns != null && wp.WritableColumns.Length > 0)
                    {
                        sb.Append("  |  Writable: ");
                        for (int i = 0; i < wp.WritableColumns.Length; i++)
                        {
                            if (i > 0) sb.Append(',');
                            sb.Append(SpreadsheetApp.Core.CellAddress.ColumnIndexToName(wp.WritableColumns[i]));
                        }
                    }
                    if (wp.InputColumnIndex != null)
                    {
                        string letter = SpreadsheetApp.Core.CellAddress.ColumnIndexToName(wp.InputColumnIndex.Value);
                        sb.Append("  |  Input: ");
                        if (!wp.AllowInputWritesForExistingRows && !wp.AllowInputWritesForEmptyRows)
                            sb.Append(letter).Append(" read-only");
                        else if (!wp.AllowInputWritesForExistingRows && wp.AllowInputWritesForEmptyRows)
                            sb.Append(letter).Append(" append-only (empty rows)");
                        else
                            sb.Append(letter).Append(" writable");
                    }
                }
            }
            catch { }
            // Schema (list a few columns)
            try
            {
                var schema = ctx.Schema;
                if (schema != null && schema.Length > 0)
                {
                    sb.Append("\r\nSchema: ");
                    int shown = 0;
                    for (int i = 0; i < schema.Length && shown < 6; i++)
                    {
                        if (shown > 0) sb.Append("; ");
                        var col = schema[i];
                        sb.Append(col.ColumnLetter);
                        if (!string.IsNullOrWhiteSpace(col.Name)) sb.Append('=').Append(col.Name);
                        shown++;
                    }
                }
            }
            catch { }
            return sb.ToString();
        }
    }
}
