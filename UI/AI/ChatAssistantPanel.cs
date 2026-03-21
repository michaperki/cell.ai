using System;
using System.Drawing;
using System.Linq;
using System.Text.Json;
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
        private readonly Action<AIPlan>? _previewPlan;
        private readonly Action? _clearPreview;
        private readonly SpreadsheetApp.Core.AI.ChatSession _session;

        private readonly TextBox _input = new() { Dock = DockStyle.Top, Multiline = true, Height = 60, BorderStyle = BorderStyle.FixedSingle };
        private readonly Button _btnPlan = new() { Text = "Plan", Dock = DockStyle.Top, Height = 32 };
        private readonly CheckBox _chkAgent = new() { Text = "Analyze selection first", Dock = DockStyle.Top, Height = 20 };
        private readonly Button _btnRevise = new() { Text = "Revise", Dock = DockStyle.Top, Height = 28 };
        private readonly Button _btnCopyObs = new() { Text = "Copy Observations", Dock = DockStyle.Top, Height = 24 };
        private readonly Button _btnHistory = new() { Text = "History…", Dock = DockStyle.Top, Height = 24 };
        private readonly Button _btnReset = new() { Text = "Reset History", Dock = DockStyle.Top, Height = 24 };
        private readonly Label _lblStatus = new() { Dock = DockStyle.Top, Height = 18, Text = string.Empty, TextAlign = ContentAlignment.MiddleLeft, ForeColor = Theme.Primary, Visible = false };
        // Advanced inspector log (hidden by default)
        private readonly RichTextBox _logBox = new() { Dock = DockStyle.Fill, ReadOnly = true, BorderStyle = BorderStyle.None, BackColor = Theme.LogBg, ForeColor = Theme.LogFg, WordWrap = true };
        private readonly Button _btnApply = new() { Text = "Apply", Dock = DockStyle.Bottom, Height = 32, Enabled = false };
        private readonly TextBox _policy = new() { Dock = DockStyle.Top, Multiline = true, Height = 48, ReadOnly = true, BorderStyle = BorderStyle.None };
        private readonly LinkLabel _policyToggle = new() { Dock = DockStyle.Top, Height = 16, Text = "\u25B6 Show advanced", TextAlign = ContentAlignment.MiddleLeft };
        private readonly CheckBox _chkHardMode = new() { Text = "Strict to selection (no out-of-bounds writes)", Dock = DockStyle.Top, Height = 20 };
        private readonly ComboBox _inputPolicy = new() { Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDownList, Height = 22 };
        private Panel _advancedPanel = null!;
        // Session context indicator
        private readonly Label _lblSession = new() { Dock = DockStyle.Top, Height = 18, Text = string.Empty, TextAlign = ContentAlignment.MiddleLeft, ForeColor = Theme.TextSecondary };
        // Mode chips
        private readonly FlowLayoutPanel _modeBar = new() { Dock = DockStyle.Top, Height = 28, FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Padding = new Padding(0, 0, 0, 4) };
        private RadioButton _modeFill = null!;
        private RadioButton _modeAppend = null!;
        private RadioButton _modeTransform = null!;
        private RadioButton _modeAsk = null!;
        private string _currentMode = "Fill"; // persisted label
        private readonly Action<string>? _onModeChanged;
        // Preview card (replaces terminal-style log by default)
        private readonly Panel _previewPanel = new() { Dock = DockStyle.Fill, BackColor = Theme.PanelBg };
        private readonly Label _previewHeader = new() { Dock = DockStyle.Top, Height = 22, Font = Theme.UISemiBold, ForeColor = Theme.TextPrimary };
        private readonly Label _previewSub = new() { Dock = DockStyle.Top, Height = 18, Font = Theme.UI, ForeColor = Theme.TextSecondary };
        private readonly Label _previewWarn = new() { Dock = DockStyle.Top, Height = 18, Font = Theme.UI, ForeColor = Theme.LogInfo };
        private readonly TableLayoutPanel _previewTable = new() { Dock = DockStyle.Fill, ColumnCount = 6, RowCount = 6, BackColor = Color.White, CellBorderStyle = TableLayoutPanelCellBorderStyle.Single, Visible = false };
        // Ask-mode thread panel (scrollable Q/A turns)
        private readonly Panel _askThreadPanel = new() { Dock = DockStyle.Fill, AutoScroll = true, BackColor = Theme.PanelBg, Visible = false };

        private Panel _inputWrapper = null!;
        private AIPlan? _currentPlan;
        private readonly Func<string, CancellationToken, Task<(AIPlan plan, string[] transcript)>>? _runAgentLoop;
        private string[] _lastTranscript = Array.Empty<string>();
        private System.Windows.Forms.Timer? _thinkingTimer;
        private int _thinkingDots;

        public ChatAssistantPanel(IChatPlanner planner, SpreadsheetApp.Core.AI.ChatSession session, Func<AIContext> getContext, Action<AIPlan> applyPlan, Func<string, CancellationToken, Task<(AIPlan plan, string[] transcript)>>? runAgentLoop = null, string? initialPrompt = null, bool autoPlan = false, Action<AIPlan>? previewPlan = null, Action? clearPreview = null, string? initialMode = null, Action<string>? onModeChanged = null)
        {
            _planner = planner;
            _session = session;
            _getContext = getContext;
            _applyPlan = applyPlan;
            _previewPlan = previewPlan;
            _clearPreview = clearPreview;
            _runAgentLoop = runAgentLoop;
            _currentMode = string.IsNullOrWhiteSpace(initialMode) ? "Fill" : initialMode!;
            _onModeChanged = onModeChanged;

            Dock = DockStyle.Fill;
            Padding = new Padding(8);
            BackColor = Theme.PanelBg;

            // ── Apply theme styles to buttons ─────────────────────
            Theme.StylePrimary(_btnPlan);
            Theme.StyleSuccess(_btnApply);
            Theme.StyleSecondary(_btnRevise);
            Theme.StyleGhost(_btnCopyObs);
            Theme.StyleDanger(_btnReset);
            Theme.StyleGhost(_btnHistory);

            // ── Style log box (advanced) ─────────────────────────
            _logBox.Font = Theme.MonoSmall;

            // ── Style policy preview (collapsible) ───────────────
            _policy.BackColor = Theme.SurfaceMuted;
            _policy.ForeColor = Theme.TextSecondary;
            _policy.Font = Theme.MonoSmall;
            _policyToggle.Font = Theme.MonoSmall;
            _policyToggle.LinkColor = Theme.TextSecondary;
            _policyToggle.ActiveLinkColor = Theme.Primary;
            _policyToggle.LinkBehavior = LinkBehavior.NeverUnderline;
            _policyToggle.Click += (_, __) => ToggleAdvanced();

            // ── Style input with focus indicator ─────────────────
            _input.Font = Theme.UI;
            _input.BorderStyle = BorderStyle.None;
            _inputWrapper = new Panel
            {
                Dock = DockStyle.Top,
                Height = _input.Height + 2,
                Padding = new Padding(1),
                BackColor = Theme.PanelBorder
            };
            _input.Dock = DockStyle.Fill;
            _input.Enter += (_, __) => _inputWrapper.BackColor = Theme.Primary;
            _input.Leave += (_, __) => _inputWrapper.BackColor = Theme.PanelBorder;
            _inputWrapper.Controls.Add(_input);

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
            _lblSession.Font = Theme.UI;

            // ── Spacers ───────────────────────────────────────────
            var spacer1 = new Label { Dock = DockStyle.Top, Height = 6, BackColor = Color.Transparent };
            var spacer2 = new Label { Dock = DockStyle.Top, Height = 4, BackColor = Color.Transparent };
            var policySep = new Label { Dock = DockStyle.Top, Height = 1, BackColor = Theme.PanelBorder };

            // ── Layout ────────────────────────────────────────────
            // Added last-to-first for Top dock (last added = topmost)
            var container = new Panel { Dock = DockStyle.Fill };
            // Preview panel (non-Ask) and Ask thread panel share the main area
            BuildPreviewPanel();
            container.Controls.Add(_askThreadPanel); // Fill — Ask mode
            container.Controls.Add(_previewPanel);   // Fill — other modes
            container.Controls.Add(_btnApply);       // Bottom
            container.Controls.Add(_btnPlan);        // Top
            container.Controls.Add(spacer1);         // Top — gap before Plan
            container.Controls.Add(_btnRevise);      // Top
            // Advanced controls are added inside _advancedPanel (hidden by default)
            _advancedPanel = new Panel { Dock = DockStyle.Top, Visible = false };
            _advancedPanel.Controls.Add(_logBox);
            var advButtons = new FlowLayoutPanel { Dock = DockStyle.Top, FlowDirection = FlowDirection.LeftToRight, Height = 28 };
            advButtons.Controls.Add(_btnCopyObs);
            advButtons.Controls.Add(_btnHistory);
            advButtons.Controls.Add(_btnReset);
            _advancedPanel.Controls.Add(advButtons);
            _advancedPanel.Controls.Add(policySep);       // separator
            _advancedPanel.Controls.Add(_policy);         // policy text
            container.Controls.Add(_advancedPanel);
            container.Controls.Add(_chkAgent);       // Top
            container.Controls.Add(_lblStatus);      // Top (provider/model)
            container.Controls.Add(_lblSession);     // Top (session status)
            container.Controls.Add(_modeBar);        // Top (mode chips)
            container.Controls.Add(spacer2);         // Top — gap before action buttons
            container.Controls.Add(_policyToggle);   // Top — advanced toggle
            container.Controls.Add(_chkHardMode);    // Top
            container.Controls.Add(_inputPolicy);    // Top
            container.Controls.Add(_inputWrapper);    // Top — topmost (wraps _input with focus border)
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
            _btnHistory.Click += (_, __) => ShowHistoryDialog();
            _btnReset.Click += (_, __) => { _session.Clear(); LogAppend("History cleared.", Theme.LogInfo); try { if (_currentMode == "Ask") RenderAskThread(); _clearPreview?.Invoke(); } catch { } };
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
                        try { _clearPreview?.Invoke(); } catch { }
                    }
                }
            };

            if (!string.IsNullOrWhiteSpace(initialPrompt)) _input.Text = initialPrompt;
            // Initialize input policy options
            try
            {
                _inputPolicy.Items.AddRange(new object[] { "Input read-only", "Append-only (empty rows)", "Writable" });
                // Mode chips
                CreateModeChips(_currentMode);
                ApplyModeUiDefaults();
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
            // Show empty state
            if (!autoPlan) ShowEmptyState();
            if (autoPlan) _ = DoPlanAsync();
        }

        private void ShowHistoryDialog()
        {
            try
            {
                var form = new Form
                {
                    Text = "Chat History",
                    StartPosition = FormStartPosition.CenterParent,
                    Size = new Size(680, 480),
                    MinimizeBox = false,
                    MaximizeBox = false,
                    FormBorderStyle = FormBorderStyle.Sizable
                };
                var list = new ListView
                {
                    Dock = DockStyle.Fill,
                    View = View.Details,
                    FullRowSelect = true,
                    GridLines = true
                };
                list.Columns.Add("Role", 100);
                list.Columns.Add("Content", 520);
                foreach (var m in _session.History)
                {
                    var it = new ListViewItem(m.Role ?? string.Empty);
                    it.SubItems.Add(m.Content ?? string.Empty);
                    list.Items.Add(it);
                }
                var btnPanel = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 40, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(8, 6, 8, 6) };
                var btnClose = new Button { Text = "Close", Width = 80 };
                var btnSave = new Button { Text = "Save JSON…", Width = 110 };
                var btnCopy = new Button { Text = "Copy JSON", Width = 100 };
                Theme.StyleGhost(btnCopy);
                Theme.StyleSecondary(btnSave);
                Theme.StylePrimary(btnClose);
                btnClose.Click += (_, __) => form.Close();
                btnCopy.Click += (_, __) =>
                {
                    try { Clipboard.SetText(SerializeHistoryJson()); } catch { }
                };
                btnSave.Click += (_, __) =>
                {
                    try
                    {
                        using var sfd = new SaveFileDialog { Filter = "JSON (*.json)|*.json|All files (*.*)|*.*", FileName = "chat_history.json" };
                        if (sfd.ShowDialog(this) == DialogResult.OK)
                        {
                            System.IO.File.WriteAllText(sfd.FileName, SerializeHistoryJson());
                        }
                    }
                    catch { }
                };
                btnPanel.Controls.AddRange(new Control[] { btnClose, btnSave, btnCopy });
                form.Controls.Add(list);
                form.Controls.Add(btnPanel);
                form.ShowDialog(this);
            }
            catch { }
        }

        private string SerializeHistoryJson()
        {
            try
            {
                var opts = new JsonSerializerOptions { WriteIndented = true };
                return JsonSerializer.Serialize(_session.History, opts);
            }
            catch { return "[]"; }
        }

        private void ShowEmptyState()
        {
            LogClear();
            LogAppend("Type a prompt and click Plan to get started.", Theme.TextMuted);
            LogAppend("", Theme.LogFg);
            LogAppend("Examples:", Theme.TextMuted);
            LogAppend("  \u2022 \"Fill column B with the uppercase version of A\"", Theme.TextMuted);
            LogAppend("  \u2022 \"Sort by column C descending\"", Theme.TextMuted);
            LogAppend("  \u2022 \"Add a SUM formula in row 11\"", Theme.TextMuted);
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

        private void LogUserBubble(string text)
        {
            if (IsDisposed || _logBox.IsDisposed) return;
            var time = DateTime.Now.ToString("HH:mm");
            _logBox.SelectionStart = _logBox.TextLength;
            _logBox.SelectionLength = 0;
            // Timestamp
            _logBox.SelectionColor = Theme.TextMuted;
            _logBox.SelectionFont = Theme.MonoSmall;
            _logBox.AppendText($"  You \u00B7 {time}\n");
            // Message body — right-indented with accent
            _logBox.SelectionColor = Color.FromArgb(147, 197, 253); // light blue
            _logBox.SelectionFont = Theme.UI;
            _logBox.AppendText($"  {text}\n\n");
            _logBox.ScrollToCaret();
        }

        private void LogAIHeader()
        {
            if (IsDisposed || _logBox.IsDisposed) return;
            var time = DateTime.Now.ToString("HH:mm");
            _logBox.SelectionStart = _logBox.TextLength;
            _logBox.SelectionLength = 0;
            _logBox.SelectionColor = Theme.TextMuted;
            _logBox.SelectionFont = Theme.MonoSmall;
            _logBox.AppendText($"  AI \u00B7 {time}\n");
            _logBox.ScrollToCaret();
        }

        private void StartThinkingAnimation()
        {
            _thinkingDots = 0;
            _lblStatus.Text = "Thinking";
            _lblStatus.Visible = true;
            _thinkingTimer?.Stop();
            _thinkingTimer?.Dispose();
            _thinkingTimer = new System.Windows.Forms.Timer { Interval = 350 };
            _thinkingTimer.Tick += (_, __) =>
            {
                _thinkingDots = (_thinkingDots % 3) + 1;
                _lblStatus.Text = "Thinking" + new string('.', _thinkingDots);
            };
            _thinkingTimer.Start();
        }

        private void StopThinkingAnimation()
        {
            _thinkingTimer?.Stop();
            _thinkingTimer?.Dispose();
            _thinkingTimer = null;
            _lblStatus.Visible = false;
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
            UpdateSessionStatus();
        }

        private async Task DoPlanAsync()
        {
            _btnPlan.Enabled = false; _btnApply.Enabled = false; _currentPlan = null;
            StartThinkingAnimation();
            LogUserBubble(_input.Text ?? string.Empty);
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
                    // Apply mode + UI policy before planning
                    ctx = ApplyUiPolicy(ctx, previewOnly: false);
                    ctx = ApplyModePolicy(ctx);
                    plan = await _planner.PlanAsync(ctx, _input.Text ?? string.Empty, cts.Token).ConfigureAwait(true);
                }
                _currentPlan = plan;
                _lastTranscript = transcript;
                LogAIHeader();
                if (useAgent && transcript.Length > 0)
                {
                    LogAppend("── Observations ──", Theme.LogObservation);
                    foreach (var line in transcript) LogAppend("  " + line, Theme.LogObservation);
                    LogAppend("── Plan ──", Theme.LogCommand);
                }
                // Build preview card and enable Apply if applicable
                int writeCount = BuildPreview(plan, ctx);
                _btnApply.Enabled = (plan.Commands.Count > 0) && _currentMode != "Ask";
                if (_btnApply.Enabled)
                {
                    try { _previewPlan?.Invoke(plan); } catch { }
                }
                // Update conversation history (keep last 10 entries)
                _session.AddUser(_input.Text ?? string.Empty);
                var asstSummary = _currentMode == "Ask" ? (plan.Answer ?? string.Empty) : string.Join("; ", plan.Commands.Select(c => c.Summarize()));
                _session.AddAssistant(asstSummary);
                if (_currentMode == "Ask") RenderAskThread();
                // Build status line with provider/model, latency, tokens, remaining context
                try
                {
                    var parts = new System.Collections.Generic.List<string>();
                    if (!string.IsNullOrWhiteSpace(plan.Provider) || !string.IsNullOrWhiteSpace(plan.Model))
                    {
                        string pv = (plan.Provider ?? string.Empty);
                        if (!string.IsNullOrWhiteSpace(plan.Model)) pv = string.IsNullOrWhiteSpace(pv) ? plan.Model! : (pv + "/" + plan.Model);
                        if (!string.IsNullOrWhiteSpace(pv)) parts.Add(pv);
                    }
                    if (plan.LatencyMs.HasValue) parts.Add(plan.LatencyMs.Value + " ms");
                    if (plan.Usage != null)
                    {
                        var u = plan.Usage;
                        string tok = $"tokens {u.InputTokens?.ToString() ?? "-"}/{u.OutputTokens?.ToString() ?? "-"}/{u.TotalTokens?.ToString() ?? "-"}";
                        parts.Add(tok);
                        if (u.RemainingContext.HasValue) parts.Add($"remaining {u.RemainingContext.Value}");
                    }
                    if (parts.Count > 0)
                    {
                        _lblStatus.Text = string.Join(" · ", parts);
                        _lblStatus.Visible = true;
                    }
                }
                catch { }
                UpdateSessionStatus();
                // Optional JSONL debug log
                try
                {
                    SpreadsheetApp.Core.AI.AILogger.Log(_chkAgent.Checked ? "agent" : "chat",
                        plan.Provider, plan.Model, plan.Usage, plan.LatencyMs,
                        _getContext()?.SheetName, _getContext()?.StartRow, _getContext()?.StartCol, _getContext()?.Rows, _getContext()?.Cols,
                        plan.RawUser, plan.RawSystem,
                        plan.Commands?.Count, writeCount);
                }
                catch { }
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
                StopThinkingAnimation();
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
            string start = (ctx.StartRow >= 0 && ctx.StartCol >= 0)
                ? SpreadsheetApp.Core.CellAddress.ToAddress(ctx.StartRow, ctx.StartCol)
                : "first empty row";
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

        private void ToggleAdvanced()
        {
            if (_advancedPanel == null) return;
            _advancedPanel.Visible = !_advancedPanel.Visible;
            _policyToggle.Text = _advancedPanel.Visible ? "\u25BC Hide advanced" : "\u25B6 Show advanced";
        }

        private void CreateModeChips(string initial)
        {
            _modeBar.Controls.Clear();
            _modeFill = NewModeChip("Fill");
            _modeAppend = NewModeChip("Append");
            _modeTransform = NewModeChip("Transform");
            _modeAsk = NewModeChip("Ask");
            _modeBar.Controls.AddRange(new Control[] { _modeFill, _modeAppend, _modeTransform, _modeAsk });
            SetMode(initial);
        }

        private RadioButton NewModeChip(string label)
        {
            var rb = new RadioButton
            {
                Appearance = Appearance.Button,
                AutoSize = true,
                Text = label,
                Padding = new Padding(10, 2, 10, 2),
                Margin = new Padding(0, 0, 8, 0),
                FlatStyle = FlatStyle.Flat
            };
            rb.CheckedChanged += (_, __) =>
            {
                if (rb.Checked) { _currentMode = label; ApplyModeUiDefaults(); _onModeChanged?.Invoke(_currentMode); }
            };
            return rb;
        }

        public void SetMode(string label)
        {
            _currentMode = label;
            _modeFill.Checked = label == "Fill";
            _modeAppend.Checked = label == "Append";
            _modeTransform.Checked = label == "Transform";
            _modeAsk.Checked = label == "Ask";
            ApplyModeUiDefaults();
        }

        private void ApplyModeUiDefaults()
        {
            // Defaults per mode and UI text emphasis
            switch (_currentMode)
            {
                case "Fill":
                    _chkHardMode.Text = "Strict to selection (no out-of-bounds writes)";
                    _chkHardMode.Checked = true;
                    _inputPolicy.SelectedIndex = 0; // Input read-only
                    _btnApply.Visible = true;
                    _previewPanel.Visible = true; _askThreadPanel.Visible = false;
                    break;
                case "Append":
                    _chkHardMode.Text = "Strict to selection (no out-of-bounds writes)";
                    _chkHardMode.Checked = true;
                    _inputPolicy.SelectedIndex = 1; // Append-only (empty rows)
                    _btnApply.Visible = true;
                    _previewPanel.Visible = true; _askThreadPanel.Visible = false;
                    break;
                case "Transform":
                    _chkHardMode.Text = "Strict to selection (no out-of-bounds writes)";
                    _chkHardMode.Checked = true;
                    _inputPolicy.SelectedIndex = 0; // Input read-only
                    _btnApply.Visible = true;
                    _previewPanel.Visible = true; _askThreadPanel.Visible = false;
                    break;
                case "Ask":
                    _chkHardMode.Text = "Context limited to current selection";
                    _chkHardMode.Checked = true; // keep on internally
                    _inputPolicy.SelectedIndex = 0;
                    _btnApply.Visible = false; // no Apply in Ask mode
                    try { _chkAgent.Checked = false; _chkAgent.Enabled = false; } catch { }
                    _previewPanel.Visible = false; _askThreadPanel.Visible = true;
                    RenderAskThread();
                    break;
            }
            if (_currentMode != "Ask") { try { _chkAgent.Enabled = true; } catch { } }
            UpdateSessionStatus();
        }

        private void UpdateSessionStatus()
        {
            try
            {
                var parts = new System.Collections.Generic.List<string>();
                if (_session.History.Count == 0) parts.Add("Fresh request");
                else parts.Add($"Using prior AI context ({_session.History.Count})");
                parts.Add("Using current selection + headers");
                _lblSession.Text = string.Join(" · ", parts);
                _lblSession.Visible = true;
            }
            catch { }
        }

        private AIContext ApplyModePolicy(AIContext ctx)
        {
            try
            {
                switch (_currentMode)
                {
                    case "Fill":
                        ctx.RequestQueriesOnly = false;
                        ctx.AnswerOnly = false;
                        ctx.AllowedCommands = new[] { "set_values", "set_formula" };
                        break;
                    case "Append":
                        ctx.RequestQueriesOnly = false;
                        ctx.AnswerOnly = false;
                        ctx.AllowedCommands = new[] { "set_values" };
                        ctx.StartRow = -1; // hint: append to first empty row within selection block
                        break;
                    case "Transform":
                        ctx.RequestQueriesOnly = false;
                        ctx.AnswerOnly = false;
                        ctx.AllowedCommands = new[] { "set_values", "set_formula", "clear_range", "transform_range" };
                        break;
                    case "Ask":
                        ctx.RequestQueriesOnly = true;
                        ctx.AnswerOnly = true;
                        ctx.AllowedCommands = Array.Empty<string>();
                        break;
                }
            }
            catch { }
            return ctx;
        }

        private void BuildPreviewPanel()
        {
            _previewPanel.Controls.Clear();
            _previewPanel.Padding = new Padding(6);
            _previewPanel.Controls.Add(_previewTable);
            _previewPanel.Controls.Add(_previewWarn);
            _previewPanel.Controls.Add(_previewSub);
            _previewPanel.Controls.Add(_previewHeader);
        }

        private int BuildPreview(AIPlan plan, AIContext ctx)
        {
            try
            {
                _previewHeader.Text = string.Empty;
                _previewSub.Text = string.Empty;
                _previewWarn.Text = string.Empty;
                _previewTable.Visible = false;
                _previewTable.Controls.Clear();
                // Ask mode: show answer card only
                if (_currentMode == "Ask")
                {
                    string ans = plan.Answer ?? string.Empty;
                    _previewHeader.Text = string.IsNullOrWhiteSpace(ans) ? "No answer" : "Answer";
                    _previewSub.Text = ans;
                    // In Ask mode, never render write tables or counts
                    return 0;
                }
                // Build aggregate summary
                int writeCount = 0;
                int cmdCount = plan.Commands.Count;
                writeCount += plan.Commands.OfType<SetValuesCommand>().Sum(c => c.Values.Length * (c.Values.Length > 0 ? c.Values[0].Length : 0));
                writeCount += plan.Commands.OfType<SetFormulaCommand>().Sum(c => c.Formulas.Length * (c.Formulas.Length > 0 ? c.Formulas[0].Length : 0));
                string modeLabel = _currentMode;
                _previewHeader.Text = (cmdCount == 0)
                    ? "No changes suggested"
                    : $"{(writeCount > 0 ? $"AI will write {writeCount} cell(s)" : "Planned actions")}";
                // Subheader: affected range (first command) and shape
                var firstWrite = plan.Commands.FirstOrDefault(c => c is SetValuesCommand || c is SetFormulaCommand);
                if (firstWrite is SetValuesCommand sv)
                {
                    _previewSub.Text = $"Target {SpreadsheetApp.Core.CellAddress.ToAddress(sv.StartRow, sv.StartCol)} · {sv.Values.Length}x{(sv.Values.Length>0?sv.Values[0].Length:0)}";
                    // Build 5x6 sample
                    RenderSampleTable(sv.Values);
                }
                else if (firstWrite is SetFormulaCommand sf)
                {
                    _previewSub.Text = $"Target {SpreadsheetApp.Core.CellAddress.ToAddress(sf.StartRow, sf.StartCol)} · {sf.Formulas.Length}x{(sf.Formulas.Length>0?sf.Formulas[0].Length:0)}";
                    RenderSampleTable(sf.Formulas);
                }
                else
                {
                    // Structural ops only — summarize
                    var summaries = plan.Commands.Take(4).Select(c => "• " + c.Summarize()).ToArray();
                    _previewSub.Text = string.Join("\r\n", summaries);
                }
                // Warnings (basic): selection hard mode enforced; non-writable columns not previewed here
                return writeCount;
            }
            catch { return 0; }
        }

        private void RenderSampleTable(string[][] values)
        {
            try
            {
                int rows = Math.Min(5, values.Length);
                int cols = Math.Min(6, values.Length > 0 ? values[0].Length : 0);
                _previewTable.ColumnCount = Math.Max(1, cols);
                _previewTable.RowCount = Math.Max(1, rows);
                _previewTable.Controls.Clear();
                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        var txt = (c < values[r].Length) ? values[r][c] : string.Empty;
                        var lab = new Label { Text = txt, AutoEllipsis = true, Dock = DockStyle.Fill, Padding = new Padding(4), Font = Theme.UI };
                        _previewTable.Controls.Add(lab, c, r);
                    }
                }
                _previewTable.Visible = rows > 0 && cols > 0;
            }
            catch { _previewTable.Visible = false; }
        }

        private void RenderAskThread()
        {
            try
            {
                _askThreadPanel.SuspendLayout();
                _askThreadPanel.Controls.Clear();
                var flow = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true, Padding = new Padding(6), BackColor = Theme.PanelBg };
                foreach (var m in _session.History)
                {
                    var isUser = string.Equals(m.Role, "user", StringComparison.OrdinalIgnoreCase);
                    var bubble = new Panel { Width = Math.Max(100, _askThreadPanel.ClientSize.Width - 32), BackColor = isUser ? Color.FromArgb(240, 247, 255) : Color.FromArgb(245, 245, 245), Padding = new Padding(8), Margin = new Padding(0, 0, 0, 8) };
                    var who = new Label { Dock = DockStyle.Top, Height = 16, Text = isUser ? "You" : "AI", Font = Theme.MonoSmall, ForeColor = Theme.TextMuted };
                    var lbl = new Label { AutoSize = false, Dock = DockStyle.Top, Text = m.Content ?? string.Empty, Font = Theme.UI, ForeColor = Theme.TextPrimary, MaximumSize = new Size(bubble.Width - 16, 0) };
                    lbl.AutoSize = true;
                    bubble.Controls.Add(lbl);
                    bubble.Controls.Add(who);
                    flow.Controls.Add(bubble);
                }
                _askThreadPanel.Controls.Add(flow);
                _askThreadPanel.ResumeLayout();
                _askThreadPanel.PerformLayout();
            }
            catch { }
        }
    }
}
