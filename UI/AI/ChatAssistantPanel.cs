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

        private readonly TextBox _input = new() { Dock = DockStyle.Fill, Multiline = true, Height = 60, BorderStyle = BorderStyle.FixedSingle };
        private readonly Button _btnPlan = new() { Text = "Send", Dock = DockStyle.Fill, Height = 32 };
        private readonly CheckBox _chkAgent = new() { Text = "Analyze selection first", Dock = DockStyle.Top, Height = 20 };
        // Revise is now per-card; remove global Revise button
        private readonly Button _btnCopyObs = new() { Text = "Copy Observations", Dock = DockStyle.Top, Height = 24 };
        private readonly Button _btnHistory = new() { Text = "History…", Dock = DockStyle.Top, Height = 24 };
        private readonly Button _btnReset = new() { Text = "Reset History", Dock = DockStyle.Top, Height = 24 };
        private readonly Label _lblStatus = new() { Dock = DockStyle.Top, Height = 18, Text = string.Empty, TextAlign = ContentAlignment.MiddleLeft, ForeColor = Theme.Primary, Visible = false };
        // Removed terminal-style log and global Apply; proposals render as inline cards
        private readonly TextBox _policy = new() { Dock = DockStyle.Top, Multiline = true, Height = 48, ReadOnly = true, BorderStyle = BorderStyle.None };
        private readonly LinkLabel _policyToggle = new() { Dock = DockStyle.Top, Height = 16, Text = "\u25B6 Show advanced", TextAlign = ContentAlignment.MiddleLeft };
        private readonly CheckBox _chkHardMode = new() { Text = "Strict to selection (no out-of-bounds writes)", Dock = DockStyle.Top, Height = 20 };
        private readonly ComboBox _inputPolicy = new() { Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDownList, Height = 22 };
        private Panel _advancedPanel = null!;
        // Session context indicator
        private readonly Label _lblSession = new() { Dock = DockStyle.Top, Height = 18, Text = string.Empty, TextAlign = ContentAlignment.MiddleLeft, ForeColor = Theme.TextSecondary };
        // Unified thread panel (scrollable conversation with inline cards)
        private readonly Panel _threadHost = new() { Dock = DockStyle.Fill, AutoScroll = true, BackColor = Theme.PanelBg };
        private FlowLayoutPanel _threadFlow = null!;

        private Panel _inputWrapper = null!;
        private Panel _composerPanel = null!;
        // Current plan per-card; no global current plan
        private readonly Func<string, CancellationToken, Task<(AIPlan plan, string[] transcript)>>? _runAgentLoop;
        private string[] _lastTranscript = Array.Empty<string>();
        private System.Windows.Forms.Timer? _thinkingTimer;
        private int _thinkingDots;
        private readonly CheckBox _chkAutoApply = new() { Text = "Auto-apply small safe writes", Dock = DockStyle.Top, Height = 18, Checked = true };
        private int _lastProposalVersion = 0;
        // Session token counters (rough usage across this panel's session)
        private long _sessionTokensIn = 0;
        private long _sessionTokensOut = 0;
        private long _sessionTokensTotal = 0;

        public ChatAssistantPanel(IChatPlanner planner, SpreadsheetApp.Core.AI.ChatSession session, Func<AIContext> getContext, Action<AIPlan> applyPlan, Func<string, CancellationToken, Task<(AIPlan plan, string[] transcript)>>? runAgentLoop = null, string? initialPrompt = null, bool autoPlan = false, Action<AIPlan>? previewPlan = null, Action? clearPreview = null, string? initialMode = null, Action<string>? onModeChanged = null)
        {
            _planner = planner;
            _session = session;
            _getContext = getContext;
            _applyPlan = applyPlan;
            _previewPlan = previewPlan;
            _clearPreview = clearPreview;
            _runAgentLoop = runAgentLoop;

            Dock = DockStyle.Fill;
            Padding = new Padding(8);
            BackColor = Theme.PanelBg;

            // ── Apply theme styles to buttons ─────────────────────
            Theme.StylePrimary(_btnPlan);
            Theme.StyleGhost(_btnCopyObs);
            Theme.StyleDanger(_btnReset);
            Theme.StyleGhost(_btnHistory);

            // ── Style log box (advanced) ─────────────────────────
            // no terminal log in unified thread UI

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
                Dock = DockStyle.Fill,
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
            // Container holds: [Top] session/status + advanced toggle/section, [Fill] thread, [Bottom] composer
            var container = new Panel { Dock = DockStyle.Fill };
            // Unified thread area (scrollable)
            BuildThreadHost();
            // Bottom composer: input + Send button
            _composerPanel = new Panel { Dock = DockStyle.Bottom, Padding = new Padding(6, 6, 6, 6) };
            var composerBorder = new Panel { Dock = DockStyle.Top, Height = 1, BackColor = Theme.PanelBorder };
            var sendHost = new Panel { Dock = DockStyle.Right, Width = 88, Padding = new Padding(6, 0, 0, 0) };
            sendHost.Controls.Add(_btnPlan); // _btnPlan.Dock = Fill
            var inputHost = new Panel { Dock = DockStyle.Fill };
            inputHost.Controls.Add(_inputWrapper);  // _inputWrapper.Dock = Fill
            _composerPanel.Controls.Add(inputHost);
            _composerPanel.Controls.Add(sendHost);
            _composerPanel.Controls.Add(composerBorder);
            // Advanced controls (collapsed by default) — move power-user toggles here
            _advancedPanel = new Panel { Dock = DockStyle.Top, Visible = false, Padding = new Padding(6, 6, 6, 6) };
            var advButtons = new FlowLayoutPanel { Dock = DockStyle.Top, FlowDirection = FlowDirection.LeftToRight, Height = 28 };
            advButtons.Controls.Add(_btnCopyObs);
            advButtons.Controls.Add(_btnHistory);
            advButtons.Controls.Add(_btnReset);
            _advancedPanel.Controls.Add(advButtons);
            _advancedPanel.Controls.Add(policySep);       // separator
            _advancedPanel.Controls.Add(_policy);         // policy text
            _advancedPanel.Controls.Add(_chkAgent);
            _advancedPanel.Controls.Add(_chkAutoApply);
            _advancedPanel.Controls.Add(_chkHardMode);
            _advancedPanel.Controls.Add(_inputPolicy);
            // Assemble container (order matters with DockStyle)
            container.Controls.Add(_threadHost);          // Fill
            container.Controls.Add(_composerPanel);       // Bottom
            container.Controls.Add(_advancedPanel);       // Top (hidden by default)
            container.Controls.Add(_policyToggle);        // Top — advanced toggle
            container.Controls.Add(_lblStatus);           // Top (provider/model)
            container.Controls.Add(_lblSession);          // Top (session status)
            Controls.Add(container);
            _btnPlan.Click += async (_, __) => await DoSendAsync();
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
            _btnReset.Click += (_, __) => { _session.Clear(); _sessionTokensIn = _sessionTokensOut = _sessionTokensTotal = 0; RenderThread(); try { _clearPreview?.Invoke(); } catch { } };

            if (!string.IsNullOrWhiteSpace(initialPrompt)) _input.Text = initialPrompt;
            // Initialize input policy options
            try
            {
                _inputPolicy.Items.AddRange(new object[] { "Input read-only", "Append-only (empty rows)", "Writable" });
                // Unified default: make follow-up edits easier by default
                _inputPolicy.SelectedIndex = 2; // Writable
                _chkHardMode.Checked = false;   // Do not fence to selection unless opted in
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
            RenderThread();
            if (autoPlan) _ = DoSendAsync();
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

        public void FocusInput()
        {
            try { _input.Focus(); } catch { }
        }

        public void SetPrompt(string text, bool autoPlan)
        {
            _input.Text = text ?? string.Empty;
            if (autoPlan) _ = DoSendAsync();
        }

        public void RefreshPolicy(AIContext ctx)
        {
            try { _policy.Text = BuildPolicyPreview(ctx); } catch { }
            UpdateSessionStatus();
        }

        public void RefreshThread()
        {
            RenderThread();
        }

        private async Task DoSendAsync()
        {
            _btnPlan.Enabled = false;
            StartThinkingAnimation();
            var userText = _input.Text ?? string.Empty;
            _session.AddUser(userText);
            RenderThread();
            try
            {
                int timeoutSec = 30;
                try { var s = Environment.GetEnvironmentVariable("AI_PLAN_TIMEOUT_SEC"); if (!string.IsNullOrWhiteSpace(s)) timeoutSec = Math.Max(5, int.Parse(s)); } catch { }
                using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(timeoutSec));
                var ctx = _getContext();
                try { _policy.Text = BuildPolicyPreview(ctx); } catch { }
                ctx.Conversation = new System.Collections.Generic.List<ChatMessage>(_session.History);
                AIPlan plan;
                string[] transcript = Array.Empty<string>();
                var run = _runAgentLoop;
                bool useAgent = _chkAgent.Checked && run != null;
                if (useAgent)
                {
                    var res = await run!(userText, cts.Token).ConfigureAwait(true);
                    plan = res.plan; transcript = res.transcript;
                }
                else
                {
                    ctx = ApplyUiPolicy(ctx, previewOnly: false);
                    ctx = ApplyIntentHeuristics(ctx, userText);
                    plan = await _planner.PlanAsync(ctx, userText, cts.Token).ConfigureAwait(true);
                }
                _lastTranscript = transcript;
                if (useAgent && transcript.Length > 0)
                {
                    foreach (var line in transcript) _session.AddObservation(line);
                }
                bool hasCommands = plan.Commands != null && plan.Commands.Count > 0;
                if (!hasCommands)
                {
                    string ans = plan.Answer ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(ans))
                    {
                        var cmds = (plan.Commands != null && plan.Commands.Count > 0) ? plan.Commands : System.Linq.Enumerable.Empty<IAICommand>();
                        ans = string.Join("; ", cmds.Select(c => c.Summarize()));
                    }
                    _session.AddAnswer(ans);
                }
                else
                {
                    int idx = _session.AddProposal(plan, ++_lastProposalVersion);
                    AddOrUpdateProposalCard(idx, plan, _lastProposalVersion);
                    if (_chkAutoApply.Checked && IsSmallSafeValuesOnly(plan, ctx, maxCells: 50))
                    {
                        try { _applyPlan(plan); _session.AddAppliedSummary(plan); }
                        catch (Exception ex) { _session.AddAnswer($"Apply error: {ex.Message}"); }
                    }
                }
                RenderThread();
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
                        var u = plan.Usage; string tok = $"tokens {u.InputTokens?.ToString() ?? "-"}/{u.OutputTokens?.ToString() ?? "-"}/{u.TotalTokens?.ToString() ?? "-"}";
                        // Accumulate into session totals
                        try { if (u.InputTokens.HasValue) _sessionTokensIn += u.InputTokens.Value; } catch { }
                        try { if (u.OutputTokens.HasValue) _sessionTokensOut += u.OutputTokens.Value; } catch { }
                        try { if (u.TotalTokens.HasValue) _sessionTokensTotal += u.TotalTokens.Value; } catch { }
                        parts.Add(tok);
                        if (u.RemainingContext.HasValue) parts.Add($"remaining {u.RemainingContext.Value}");
                    }
                    if (parts.Count > 0) { _lblStatus.Text = string.Join(" · ", parts); _lblStatus.Visible = true; }
                }
                catch { }
                UpdateSessionStatus();
                try
                {
                    SpreadsheetApp.Core.AI.AILogger.Log(_chkAgent.Checked ? "agent" : "chat",
                        plan.Provider, plan.Model, plan.Usage, plan.LatencyMs,
                        _getContext()?.SheetName, _getContext()?.StartRow, _getContext()?.StartCol, _getContext()?.Rows, _getContext()?.Cols,
                        plan.RawUser, plan.RawSystem,
                        plan.Commands?.Count, (plan.Commands?.OfType<SetValuesCommand>().Sum(c => c.Values.Length * (c.Values.FirstOrDefault()?.Length ?? 0)) ?? 0));
                }
                catch { }
            }
            catch (OperationCanceledException)
            {
                _session.AddAnswer("Planning canceled or timed out.");
                RenderThread();
            }
            catch (Exception ex)
            {
                _session.AddAnswer($"Error: {ex.Message}");
                RenderThread();
            }
            finally
            {
                _btnPlan.Enabled = true;
                StopThinkingAnimation();
                _input.Clear();
                _input.Focus();
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

        // Mode chips removed in unified UI

        private void UpdateSessionStatus()
        {
            try
            {
                var parts = new System.Collections.Generic.List<string>();
                if (_session.History.Count == 0) parts.Add("Fresh request");
                else parts.Add($"Using prior AI context ({_session.History.Count})");
                parts.Add("Using current selection + headers");
                if (_sessionTokensTotal > 0)
                {
                    parts.Add($"Session tokens {_sessionTokensIn}/{_sessionTokensOut}/{_sessionTokensTotal}");
                }
                _lblSession.Text = string.Join(" · ", parts);
                _lblSession.Visible = true;
            }
            catch { }
        }

        // No mode policy in unified UI; heuristics are applied in ApplyIntentHeuristics

        // Preview helpers removed in unified UI (cards render their own samples)

        private void RenderAskThread()
        {
            // unused in unified UI
        }

        private void BuildThreadHost()
        {
            _threadHost.Controls.Clear();
            _threadFlow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true,
                Padding = new Padding(6),
                BackColor = Theme.PanelBg
            };
            _threadHost.Controls.Add(_threadFlow);
        }

        private void RenderThread()
        {
            try
            {
                _threadFlow.SuspendLayout();
                _threadFlow.Controls.Clear();
                int width = Math.Max(100, _threadHost.ClientSize.Width - 16);
                foreach (var e in _session.Thread)
                {
                    switch (e.Type)
                    {
                        case SpreadsheetApp.Core.AI.ChatEntryType.User:
                        {
                            var bubble = new Panel { Width = width, BackColor = Color.FromArgb(240, 247, 255), Padding = new Padding(6), Margin = new Padding(0, 0, 0, 6) };
                            var who = new Label { Dock = DockStyle.Top, Height = 14, Text = "You", Font = Theme.MonoSmall, ForeColor = Theme.TextMuted };
                            var lbl = new Label { AutoSize = true, Dock = DockStyle.Top, Text = e.Content ?? string.Empty, Font = Theme.UI, ForeColor = Theme.TextPrimary, MaximumSize = new Size(width - 12, 0) };
                            bubble.Controls.Add(lbl); bubble.Controls.Add(who); _threadFlow.Controls.Add(bubble);
                            break;
                        }
                        case SpreadsheetApp.Core.AI.ChatEntryType.Answer:
                        {
                            var bubble = new Panel { Width = width, BackColor = Color.FromArgb(245, 245, 245), Padding = new Padding(6), Margin = new Padding(0, 0, 0, 6) };
                            var who = new Label { Dock = DockStyle.Top, Height = 14, Text = "AI", Font = Theme.MonoSmall, ForeColor = Theme.TextMuted };
                            var lbl = new Label { AutoSize = true, Dock = DockStyle.Top, Text = e.Content ?? string.Empty, Font = Theme.UI, ForeColor = Theme.TextPrimary, MaximumSize = new Size(width - 12, 0) };
                            bubble.Controls.Add(lbl); bubble.Controls.Add(who); _threadFlow.Controls.Add(bubble);
                            break;
                        }
                        case SpreadsheetApp.Core.AI.ChatEntryType.Observation:
                        {
                            var lbl = new Label { Width = width, AutoSize = true, Text = e.Content ?? string.Empty, Font = Theme.MonoSmall, ForeColor = Theme.LogObservation, Margin = new Padding(0, 0, 0, 2) };
                            _threadFlow.Controls.Add(lbl);
                            break;
                        }
                        case SpreadsheetApp.Core.AI.ChatEntryType.ActionProposal:
                        {
                            var card = new ActionCardControl();
                            card.Width = width; card.SetPlan(e.Proposal ?? new AIPlan(), e.ProposalVersion);
                            card.ApplyRequested += (_, __) => { try { _applyPlan(card.GetPlan()); _session.AddAppliedSummary(card.GetPlan()); RenderThread(); } catch (Exception ex) { _session.AddAnswer($"Apply error: {ex.Message}"); RenderThread(); } };
                            card.ReviseRequested += async (_, __) => await ReviseProposalAsync(e, card);
                            card.CardFocused += (_, __) => { try { _previewPlan?.Invoke(card.GetPlan()); } catch { } };
                            _threadFlow.Controls.Add(card);
                            break;
                        }
                        case SpreadsheetApp.Core.AI.ChatEntryType.AppliedSummary:
                        {
                            var lbl = new Label { Width = width, AutoSize = true, Text = $"Applied: {e.Content}", Font = Theme.UI, ForeColor = Theme.LogSuccess, Margin = new Padding(0, 0, 0, 8) };
                            _threadFlow.Controls.Add(lbl);
                            break;
                        }
                    }
                }
                _threadFlow.ResumeLayout();
                _threadFlow.PerformLayout();
                // Auto-scroll to the most recent entry so it appears above the composer
                try
                {
                    if (_threadFlow.Controls.Count > 0)
                    {
                        var last = _threadFlow.Controls[_threadFlow.Controls.Count - 1];
                        _threadFlow.ScrollControlIntoView(last);
                    }
                }
                catch { }
            }
            catch { }
        }

        private void AddOrUpdateProposalCard(int entryIndex, AIPlan plan, int version)
        {
            // Simple path: re-render thread and trigger preview
            RenderThread();
            try { _previewPlan?.Invoke(plan); } catch { }
        }

        private async System.Threading.Tasks.Task ReviseProposalAsync(SpreadsheetApp.Core.AI.ChatEntry entry, ActionCardControl card)
        {
            try
            {
                string feedback = _input.Text ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(feedback)) _session.AddUser(feedback);
                RenderThread();
                int timeoutSec = 30; try { var s = Environment.GetEnvironmentVariable("AI_PLAN_TIMEOUT_SEC"); if (!string.IsNullOrWhiteSpace(s)) timeoutSec = Math.Max(5, int.Parse(s)); } catch { }
                using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(timeoutSec));
                var ctx = ApplyIntentHeuristics(ApplyUiPolicy(_getContext(), previewOnly: false), feedback);
                ctx.Conversation = new System.Collections.Generic.List<ChatMessage>(_session.History);
                var newPlan = await _planner.PlanAsync(ctx, feedback, cts.Token).ConfigureAwait(true);
                int newVersion = entry.ProposalVersion + 1;
                card.SetPlan(newPlan, newVersion);
                var cmds = (newPlan.Commands != null && newPlan.Commands.Count > 0) ? newPlan.Commands : System.Linq.Enumerable.Empty<IAICommand>();
                entry.Proposal = newPlan; entry.ProposalVersion = newVersion; entry.Content = string.Join("; ", cmds.Select(c => c.Summarize()));
                RenderThread();
            }
            catch (Exception ex)
            {
                _session.AddAnswer($"Revise failed: {ex.Message}");
                RenderThread();
            }
            finally
            {
                try { _clearPreview?.Invoke(); } catch { }
                _input.Clear(); _input.Focus();
            }
        }

        private static bool IsSmallSafeValuesOnly(AIPlan plan, AIContext ctx, int maxCells)
        {
            try
            {
                if (plan.Commands == null || plan.Commands.Count == 0) return false;
                if (plan.Commands.Any(c => c is not SetValuesCommand)) return false; // values-only
                int total = plan.Commands.OfType<SetValuesCommand>().Sum(c => c.Values.Length * (c.Values.FirstOrDefault()?.Length ?? 0));
                if (total <= 0 || total > maxCells) return false;
                // within selection bounds
                int r0 = Math.Max(0, ctx.StartRow); int c0 = Math.Max(0, ctx.StartCol);
                int r1 = r0 + Math.Max(1, ctx.Rows) - 1; int c1 = c0 + Math.Max(1, ctx.Cols) - 1;
                foreach (var sv in plan.Commands.OfType<SetValuesCommand>())
                {
                    int rows = sv.Values.Length; int cols = rows > 0 ? sv.Values[0].Length : 0;
                    int sr0 = sv.StartRow; int sc0 = sv.StartCol; int sr1 = sr0 + rows - 1; int sc1 = sc0 + cols - 1;
                    if (sr0 < r0 || sc0 < c0 || sr1 > r1 || sc1 > c1) return false;
                }
                return true;
            }
            catch { return false; }
        }

        private static AIContext ApplyIntentHeuristics(AIContext ctx, string prompt)
        {
            try
            {
                var p = (prompt ?? string.Empty).ToLowerInvariant();
                ctx.RequestQueriesOnly = false; ctx.AnswerOnly = false;
                if (p.Contains("insert column") || p.Contains("add column") || p.Contains("create column") || p.Contains("new column"))
                {
                    ctx.AllowedCommands = new[] { "insert_cols" };
                    try { var wp = ctx.WritePolicy ?? new SelectionWritePolicy(); wp.HeaderRowReadOnly = false; ctx.WritePolicy = wp; } catch { }
                }
                else if (p.Contains("delete column")) ctx.AllowedCommands = new[] { "delete_cols" };
                else if (p.Contains("insert row")) ctx.AllowedCommands = new[] { "insert_rows" };
                else if (p.Contains("delete row")) ctx.AllowedCommands = new[] { "delete_rows" };
                else if (p.Contains("copy ")) ctx.AllowedCommands = new[] { "copy_range" };
                else if (p.Contains("move ")) ctx.AllowedCommands = new[] { "move_range" };
                else if (p.Contains("rename sheet")) ctx.AllowedCommands = new[] { "rename_sheet" };
                else if (p.Contains("clear ")) ctx.AllowedCommands = new[] { "clear_range" };
                else if (p.Contains("format")) ctx.AllowedCommands = new[] { "set_format" };
                else if (p.Contains("validation")) ctx.AllowedCommands = new[] { "set_validation" };
                else if (p.Contains("conditional format") || p.Contains("condfmt")) ctx.AllowedCommands = new[] { "set_conditional_format" };
                else if (p.Contains("transform ") || p.Contains("cleanup")) ctx.AllowedCommands = new[] { "transform_range" };
                else if (p.Contains("sort by") || p.Contains("sort ")) ctx.AllowedCommands = new[] { "sort_range" };
                else if (p.Contains("explain") || p.Contains("what is") || p.Contains("why")) { ctx.AllowedCommands = Array.Empty<string>(); ctx.AnswerOnly = true; }
                else ctx.AllowedCommands = null;
            }
            catch { }
            return ctx;
        }
    }
}
