using System;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using SpreadsheetApp.Core.AI;
using SpreadsheetApp.UI;

namespace SpreadsheetApp.UI.AI
{
    // Dockable chat assistant panel (reuses ChatAssistantForm logic in a UserControl)
    public sealed class ChatAssistantPanel : UserControl
    {
        private readonly IChatPlanner _planner;
        private readonly Func<AIContext> _getContext;
        private readonly Action<AIPlan> _applyPlan;
        private readonly SpreadsheetApp.Core.AI.ChatSession _session;

        private readonly TextBox _input = new() { Dock = DockStyle.Top, Multiline = true, Height = 60, BorderStyle = BorderStyle.FixedSingle };
        private readonly Button _btnPlan = new() { Text = "Plan", Dock = DockStyle.Top, Height = 32 };
        private readonly CheckBox _chkAgent = new() { Text = "Let AI explore first", Dock = DockStyle.Top, Height = 20 };
        private readonly Button _btnRevise = new() { Text = "Revise", Dock = DockStyle.Top, Height = 28 };
        private readonly Button _btnCopyObs = new() { Text = "Copy Observations", Dock = DockStyle.Top, Height = 24 };
        private readonly Button _btnReset = new() { Text = "Reset History", Dock = DockStyle.Top, Height = 24 };
        private readonly Label _lblStatus = new() { Dock = DockStyle.Top, Height = 18, Text = string.Empty, TextAlign = ContentAlignment.MiddleLeft, ForeColor = Theme.Primary, Visible = false };
        private readonly RichTextBox _logBox = new() { Dock = DockStyle.Fill, ReadOnly = true, BorderStyle = BorderStyle.None, BackColor = Theme.LogBg, ForeColor = Theme.LogFg, WordWrap = true };
        private readonly Button _btnApply = new() { Text = "Apply", Dock = DockStyle.Bottom, Height = 32, Enabled = false };
        private readonly TextBox _policy = new() { Dock = DockStyle.Top, Multiline = true, Height = 48, ReadOnly = true, BorderStyle = BorderStyle.None };
        private readonly CheckBox _chkHardMode = new() { Text = "Selection hard mode (no out-of-bounds writes)", Dock = DockStyle.Top, Height = 20 };
        private readonly ComboBox _inputPolicy = new() { Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDownList, Height = 22 };

        private AIPlan? _currentPlan;
        private readonly Func<string, CancellationToken, Task<(AIPlan plan, string[] transcript)>>? _runAgentLoop;
        private string[] _lastTranscript = Array.Empty<string>();

        public ChatAssistantPanel(IChatPlanner planner, SpreadsheetApp.Core.AI.ChatSession session, Func<AIContext> getContext, Action<AIPlan> applyPlan, Func<string, CancellationToken, Task<(AIPlan plan, string[] transcript)>>? runAgentLoop = null, string? initialPrompt = null, bool autoPlan = false)
        {
            _planner = planner;
            _session = session;
            _getContext = getContext;
            _applyPlan = applyPlan;
            _runAgentLoop = runAgentLoop;

            Dock = DockStyle.Fill;
            Padding = new Padding(8);
            BackColor = Theme.PanelBg;

            // ── Apply theme styles to buttons ─────────────────────
            Theme.StylePrimary(_btnPlan);
            Theme.StyleSuccess(_btnApply);
            Theme.StyleSecondary(_btnRevise);
            Theme.StyleGhost(_btnCopyObs);
            Theme.StyleDanger(_btnReset);

            // ── Style log box ─────────────────────────────────────
            _logBox.Font = Theme.MonoSmall;

            // ── Style policy preview ──────────────────────────────
            _policy.BackColor = Theme.SurfaceMuted;
            _policy.ForeColor = Theme.TextSecondary;
            _policy.Font = Theme.MonoSmall;

            // ── Style input ───────────────────────────────────────
            _input.Font = Theme.UI;

            // ── Style checkboxes ──────────────────────────────────
            _chkAgent.Font = Theme.UI;
            _chkAgent.ForeColor = Theme.TextSecondary;
            _chkHardMode.Font = Theme.UI;
            _chkHardMode.ForeColor = Theme.TextSecondary;

            // ── Style combo ───────────────────────────────────────
            _inputPolicy.Font = Theme.UI;
            _inputPolicy.FlatStyle = FlatStyle.Flat;

            // ── Status label ──────────────────────────────────────
            _lblStatus.Font = Theme.UISemiBold;

            // ── Spacers ───────────────────────────────────────────
            var spacer1 = new Label { Dock = DockStyle.Top, Height = 6, BackColor = Color.Transparent };
            var spacer2 = new Label { Dock = DockStyle.Top, Height = 4, BackColor = Color.Transparent };
            var policySep = new Label { Dock = DockStyle.Top, Height = 1, BackColor = Theme.PanelBorder };

            // ── Layout ────────────────────────────────────────────
            // Added last-to-first for Top dock (last added = topmost)
            var container = new Panel { Dock = DockStyle.Fill };
            container.Controls.Add(_logBox);         // Fill — center
            container.Controls.Add(_btnApply);       // Bottom
            container.Controls.Add(_btnPlan);        // Top
            container.Controls.Add(spacer1);         // Top — gap before Plan
            container.Controls.Add(_btnRevise);      // Top
            container.Controls.Add(_btnCopyObs);     // Top
            container.Controls.Add(_chkAgent);       // Top
            container.Controls.Add(_lblStatus);      // Top
            container.Controls.Add(spacer2);         // Top — gap before action buttons
            container.Controls.Add(_btnReset);       // Top
            container.Controls.Add(policySep);       // Top — 1px line under policy
            container.Controls.Add(_policy);         // Top
            container.Controls.Add(_chkHardMode);    // Top
            container.Controls.Add(_inputPolicy);    // Top
            container.Controls.Add(_input);          // Top — topmost
            Controls.Add(container);

            _btnPlan.Click += async (_, __) => await DoPlanAsync();
            _btnRevise.Click += async (_, __) =>
            {
                var feedback = _input.Text ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(feedback))
                {
                    _session.AddUser(feedback);
                }
                await DoPlanAsync();
            };
            _btnCopyObs.Click += (_, __) =>
            {
                try
                {
                    if (_lastTranscript != null && _lastTranscript.Length > 0)
                    {
                        Clipboard.SetText(string.Join("\r\n", _lastTranscript));
                    }
                }
                catch { }
            };
            _btnReset.Click += (_, __) => { _session.Clear(); LogAppend("History cleared.", Theme.LogInfo); };
            _btnApply.Click += (_, __) =>
            {
                if (_currentPlan != null)
                {
                    try
                    {
                        _applyPlan(_currentPlan);
                        var summary = string.Join("; ", _currentPlan.Commands.Select(c => c.Summarize()));
                        LogAppend($"Applied: {summary}", Theme.LogSuccess);
                    }
                    catch (Exception ex)
                    {
                        LogAppend($"Apply error: {ex.Message}", Theme.LogError);
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
            // Initialize input policy options
            try
            {
                _inputPolicy.Items.AddRange(new object[] { "Input read-only", "Append-only (empty rows)", "Writable" });
                _inputPolicy.SelectedIndex = 1; // default append-only
                _inputPolicy.SelectedIndexChanged += (_, __) =>
                {
                    try { var ctx = _getContext(); _policy.Text = BuildPolicyPreview(ApplyUiPolicy(ctx, previewOnly: true)); } catch { }
                };
                _chkHardMode.CheckedChanged += (_, __) =>
                {
                    try { var ctx = _getContext(); ctx.SelectionHardMode = _chkHardMode.Checked; _policy.Text = BuildPolicyPreview(ctx); } catch { }
                };
            }
            catch { }
            try { var ctx0 = _getContext(); _policy.Text = BuildPolicyPreview(ctx0); } catch { }
            if (autoPlan) _ = DoPlanAsync();
        }

        // ── Colored log helpers ──────────────────────────────────

        private void LogAppend(string text, Color color)
        {
            if (IsDisposed || _logBox.IsDisposed) return;
            _logBox.SelectionStart = _logBox.TextLength;
            _logBox.SelectionLength = 0;
            _logBox.SelectionColor = color;
            _logBox.AppendText(text + "\n");
            _logBox.ScrollToCaret();
        }

        private void LogClear()
        {
            if (IsDisposed || _logBox.IsDisposed) return;
            _logBox.Clear();
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

        public void RefreshPolicy(AIContext ctx)
        {
            try { _policy.Text = BuildPolicyPreview(ctx); } catch { }
        }

        private async Task DoPlanAsync()
        {
            _btnPlan.Enabled = false; _btnApply.Enabled = false; _currentPlan = null;
            _lblStatus.Text = "Thinking\u2026"; _lblStatus.Visible = true;
            LogAppend("Planning\u2026", Theme.LogInfo);
            try
            {
                int timeoutSec = 30;
                try { var s = Environment.GetEnvironmentVariable("AI_PLAN_TIMEOUT_SEC"); if (!string.IsNullOrWhiteSpace(s)) timeoutSec = Math.Max(5, int.Parse(s)); } catch { }
                using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(timeoutSec));
                var ctx = _getContext();
                try { _policy.Text = BuildPolicyPreview(ctx); } catch { }
                // Include rolling conversation
                ctx.Conversation = new System.Collections.Generic.List<ChatMessage>(_session.History);
                AIPlan plan;
                string[] transcript = Array.Empty<string>();
                var run = _runAgentLoop; // local for nullability analysis
                bool useAgent = _chkAgent.Checked && run != null;
                if (useAgent)
                {
                    var res = await run!(_input.Text ?? string.Empty, cts.Token).ConfigureAwait(true);
                    plan = res.plan; transcript = res.transcript;
                }
                else
                {
                    // Apply UI policy toggles before planning
                    ctx = ApplyUiPolicy(ctx, previewOnly: false);
                    plan = await _planner.PlanAsync(ctx, _input.Text ?? string.Empty, cts.Token).ConfigureAwait(true);
                }
                _currentPlan = plan;
                LogClear();
                _lastTranscript = transcript;
                if (useAgent && transcript.Length > 0)
                {
                    LogAppend("── Observations ──", Theme.LogObservation);
                    foreach (var line in transcript) LogAppend("  " + line, Theme.LogObservation);
                    LogAppend("── Plan ──", Theme.LogCommand);
                }
                if (plan.Commands.Count == 0)
                {
                    LogAppend("No changes suggested.", Theme.LogInfo);
                }
                else
                {
                    foreach (var cmd in plan.Commands)
                    {
                        LogAppend(cmd.Summarize(), Theme.LogCommand);
                        string? rationale = TryGetRationale(cmd);
                        if (!string.IsNullOrWhiteSpace(rationale)) LogAppend("  \u2192 " + rationale, Theme.LogRationale);
                    }
                    int writeCount = 0;
                    writeCount += plan.Commands.OfType<SetValuesCommand>().Sum(c => c.Values.Length * (c.Values.Length > 0 ? c.Values[0].Length : 0));
                    writeCount += plan.Commands.OfType<SetFormulaCommand>().Sum(c => c.Formulas.Length * (c.Formulas.Length > 0 ? c.Formulas[0].Length : 0));
                    LogAppend($"Total writes: {writeCount}", Theme.LogInfo);
                    _btnApply.Enabled = true;
                }
                // Update conversation history (keep last 10 entries)
                _session.AddUser(_input.Text ?? string.Empty);
                var asstSummary = string.Join("; ", plan.Commands.Select(c => c.Summarize()));
                _session.AddAssistant(asstSummary);
            }
            catch (OperationCanceledException)
            {
                LogAppend("Planning canceled or timed out.", Theme.LogError);
            }
            catch (Exception ex)
            {
                LogAppend($"Error: {ex.Message}", Theme.LogError);
            }
            finally
            {
                _btnPlan.Enabled = true;
                _lblStatus.Visible = false;
            }
        }

        private static string? TryGetRationale(IAICommand cmd)
        {
            try
            {
                switch (cmd)
                {
                    case SetValuesCommand sv: return sv.Rationale;
                    case SetFormulaCommand sf: return sf.Rationale;
                    case SetTitleCommand st: return st.Rationale;
                    case ClearRangeCommand cr: return cr.Rationale;
                    case SortRangeCommand sr: return sr.Rationale;
                    case CreateSheetCommand cs: return cs.Rationale;
                    case RenameSheetCommand rn: return rn.Rationale;
                    case InsertRowsCommand ir: return ir.Rationale;
                    case DeleteRowsCommand dr: return dr.Rationale;
                    case InsertColsCommand ic: return ic.Rationale;
                    case DeleteColsCommand dc: return dc.Rationale;
                    case DeleteSheetCommand ds: return ds.Rationale;
                    case CopyRangeCommand cp: return cp.Rationale;
                    case MoveRangeCommand mv: return mv.Rationale;
                    case SetFormatCommand fm: return fm.Rationale;
                    case SetValidationCommand svv: return svv.Rationale;
                    case SetConditionalFormatCommand scf: return scf.Rationale;
                    case TransformRangeCommand tr: return tr.Rationale;
                    default: return null;
                }
            }
            catch { return null; }
        }

        private AIContext ApplyUiPolicy(AIContext ctx, bool previewOnly)
        {
            try
            {
                var wp = ctx.WritePolicy ?? new SelectionWritePolicy();
                switch (_inputPolicy.SelectedIndex)
                {
                    case 0: // read-only
                        wp.AllowInputWritesForExistingRows = false;
                        wp.AllowInputWritesForEmptyRows = false;
                        break;
                    case 1: // append-only
                        wp.AllowInputWritesForExistingRows = false;
                        wp.AllowInputWritesForEmptyRows = true;
                        break;
                    case 2: // writable
                        wp.AllowInputWritesForExistingRows = true;
                        wp.AllowInputWritesForEmptyRows = true;
                        break;
                    default:
                        break;
                }
                ctx.WritePolicy = wp;
                ctx.SelectionHardMode = _chkHardMode.Checked;
            }
            catch { }
            return ctx;
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
