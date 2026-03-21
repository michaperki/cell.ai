using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using SpreadsheetApp.Core.AI;

namespace SpreadsheetApp.UI.AI
{
    // Inline action card rendering a proposed AIPlan with Apply/Revise actions
    public sealed class ActionCardControl : UserControl
    {
        public event EventHandler? ApplyRequested;
        public event EventHandler? ReviseRequested;
        public event EventHandler? CardFocused; // host may preview overlay on focus/hover

        private AIPlan _plan = new AIPlan();
        private readonly Label _title = new() { Dock = DockStyle.Top, Height = 22, Font = Theme.UISemiBold, ForeColor = Theme.TextPrimary };
        private readonly Label _meta = new() { Dock = DockStyle.Top, Height = 16, Font = Theme.MonoSmall, ForeColor = Theme.TextSecondary };
        private readonly Label _summary = new() { Dock = DockStyle.Top, AutoSize = false, Height = 40, Font = Theme.UI, ForeColor = Theme.TextPrimary };
        private readonly TableLayoutPanel _sample = new() { Dock = DockStyle.Top, ColumnCount = 1, RowCount = 1, BackColor = Color.White, CellBorderStyle = TableLayoutPanelCellBorderStyle.Single, Visible = false, Height = 120 };
        private readonly FlowLayoutPanel _buttons = new() { Dock = DockStyle.Top, FlowDirection = FlowDirection.LeftToRight, Height = 36, Padding = new Padding(0, 6, 0, 6) };
        private readonly Button _btnApply = new() { Text = "Apply", Width = 80 };
        private readonly Button _btnRevise = new() { Text = "Revise", Width = 80 };

        public int Version { get; private set; } = 1;

        public ActionCardControl()
        {
            BackColor = Theme.PanelBg;
            Padding = new Padding(6);
            Theme.StylePrimary(_btnApply);
            Theme.StyleGhost(_btnRevise);
            _buttons.Controls.Add(_btnApply);
            _buttons.Controls.Add(_btnRevise);
            Controls.Add(_buttons);
            Controls.Add(_sample);
            Controls.Add(_summary);
            Controls.Add(_meta);
            Controls.Add(_title);
            Height = 200;

            _btnApply.Click += (_, __) => ApplyRequested?.Invoke(this, EventArgs.Empty);
            _btnRevise.Click += (_, __) => ReviseRequested?.Invoke(this, EventArgs.Empty);
            GotFocus += (_, __) => CardFocused?.Invoke(this, EventArgs.Empty);
            MouseEnter += (_, __) => CardFocused?.Invoke(this, EventArgs.Empty);
        }

        public void SetPlan(AIPlan plan, int version)
        {
            _plan = plan ?? new AIPlan();
            Version = version;
            Build();
        }

        private void Build()
        {
            try
            {
                _title.Text = Version > 1 ? $"Proposed changes (v{Version})" : "Proposed changes";
                string prov = string.Empty;
                if (!string.IsNullOrWhiteSpace(_plan.Provider) || !string.IsNullOrWhiteSpace(_plan.Model))
                {
                    prov = string.IsNullOrWhiteSpace(_plan.Provider) ? _plan.Model ?? string.Empty : ($"{_plan.Provider}/{_plan.Model}");
                }
                var usage = _plan.Usage;
                string usageTxt = (usage != null) ? ($"tokens {usage.InputTokens?.ToString() ?? "-"}/{usage.OutputTokens?.ToString() ?? "-"}/{usage.TotalTokens?.ToString() ?? "-"}") : string.Empty;
                string lat = _plan.LatencyMs.HasValue ? ($"{_plan.LatencyMs.Value} ms") : string.Empty;
                _meta.Text = string.Join(" · ", new[] { prov, lat, usageTxt }.Where(s => !string.IsNullOrWhiteSpace(s)));

                // Summary
                string summary = string.Empty;
                try
                {
                    if (_plan.Commands != null && _plan.Commands.Count > 0)
                    {
                        var parts = _plan.Commands.Take(4).Select(c => "• " + c.Summarize());
                        summary = string.Join("\r\n", parts);
                    }
                }
                catch { }
                _summary.Text = summary;

                // Build sample for first write command
                _sample.Visible = false;
                _sample.Controls.Clear();
                var firstWrite = _plan.Commands?.FirstOrDefault(c => c is SetValuesCommand || c is SetFormulaCommand);
                if (firstWrite is SetValuesCommand sv)
                {
                    RenderSample(sv.Values);
                }
                else if (firstWrite is SetFormulaCommand sf)
                {
                    RenderSample(sf.Formulas);
                }
            }
            catch { }
        }

        private void RenderSample(string[][] values)
        {
            try
            {
                int rows = Math.Min(5, values.Length);
                int cols = Math.Min(6, values.Length > 0 ? values[0].Length : 0);
                if (rows <= 0 || cols <= 0) { _sample.Visible = false; return; }
                _sample.RowStyles.Clear();
                _sample.ColumnStyles.Clear();
                _sample.ColumnCount = cols; _sample.RowCount = rows;
                _sample.Controls.Clear();
                for (int c = 0; c < cols; c++) _sample.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f / cols));
                for (int r = 0; r < rows; r++) _sample.RowStyles.Add(new RowStyle(SizeType.Percent, 100f / rows));
                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        var txt = (c < values[r].Length) ? values[r][c] : string.Empty;
                        var lab = new Label { Text = txt, AutoEllipsis = true, Dock = DockStyle.Fill, Padding = new Padding(4), Font = Theme.UI };
                        _sample.Controls.Add(lab, c, r);
                    }
                }
                _sample.Visible = true;
            }
            catch { _sample.Visible = false; }
        }

        public AIPlan GetPlan() => _plan;
    }
}

