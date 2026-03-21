using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using SpreadsheetApp.Core.AI;
using SpreadsheetApp.UI;

namespace SpreadsheetApp.UI.AI
{
    public class GenerateFillDialog : Form
    {
        private readonly IInferenceProvider _provider;
        private readonly AIContext _baseContext;
        private readonly TextBox _txtPrompt = new() { Multiline = true, Height = 60, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
        private readonly NumericUpDown _numRows = new() { Minimum = 1, Maximum = 1000, Value = 10, Width = 80 };
        private readonly NumericUpDown _numCols = new() { Minimum = 1, Maximum = 50, Value = 1, Width = 80 };
        private readonly Button _btnPreview = new() { Text = "Preview" };
        private readonly Button _btnAccept = new() { Text = "Accept", Enabled = false };
        private readonly Button _btnCancel = new() { Text = "Cancel" };
        private readonly DataGridView _preview = new() { ReadOnly = true, AllowUserToAddRows = false, AllowUserToDeleteRows = false, Dock = DockStyle.Fill };
        private CancellationTokenSource? _cts;

        public string[][] ResultCells { get; private set; } = Array.Empty<string[]>();

        public GenerateFillDialog(IInferenceProvider provider, AIContext baseContext)
        {
            _provider = provider;
            _baseContext = baseContext;
            Text = "AI Generate Fill";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimizeBox = false; MaximizeBox = false;
            Width = 640; Height = 480;
            BackColor = System.Drawing.Color.White;
            Font = Theme.UI;

            Theme.StyleSecondary(_btnPreview);
            Theme.StylePrimary(_btnAccept);
            Theme.StyleGhost(_btnCancel);
            Theme.StyleGrid(_preview);

            var pnlTop = new TableLayoutPanel { Dock = DockStyle.Top, Height = 120, ColumnCount = 6, RowCount = 2, AutoSize = false };
            pnlTop.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            pnlTop.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            pnlTop.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            pnlTop.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            pnlTop.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            pnlTop.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            var lblPrompt = new Label { Text = "Prompt:", AutoSize = true, Anchor = AnchorStyles.Left, ForeColor = Theme.TextPrimary };
            var lblRows = new Label { Text = "Rows:", AutoSize = true, Anchor = AnchorStyles.Right, ForeColor = Theme.TextPrimary };
            var lblCols = new Label { Text = "Cols:", AutoSize = true, Anchor = AnchorStyles.Right, ForeColor = Theme.TextPrimary };

            pnlTop.Controls.Add(lblPrompt, 0, 0);
            pnlTop.Controls.Add(_txtPrompt, 1, 0);
            pnlTop.SetColumnSpan(_txtPrompt, 5);
            pnlTop.Controls.Add(lblRows, 2, 1);
            pnlTop.Controls.Add(_numRows, 3, 1);
            pnlTop.Controls.Add(lblCols, 4, 1);
            pnlTop.Controls.Add(_numCols, 5, 1);

            var pnlButtons = new FlowLayoutPanel { Dock = DockStyle.Bottom, FlowDirection = FlowDirection.RightToLeft, Height = 40 };
            pnlButtons.Controls.AddRange(new Control[] { _btnCancel, _btnAccept, _btnPreview });

            _btnPreview.Click += async (_, __) => await DoPreviewAsync();
            _btnAccept.Click += (_, __) => { DialogResult = DialogResult.OK; Close(); };
            _btnCancel.Click += (_, __) => { DialogResult = DialogResult.Cancel; Close(); };

            Controls.Add(_preview);
            Controls.Add(pnlButtons);
            Controls.Add(pnlTop);

            // Initialize size pickers from the provided base context
            try
            {
                var rows = Math.Max((int)_numRows.Minimum, Math.Min((int)_numRows.Maximum, _baseContext.Rows));
                var cols = Math.Max((int)_numCols.Minimum, Math.Min((int)_numCols.Maximum, _baseContext.Cols));
                _numRows.Value = rows;
                _numCols.Value = cols;
            }
            catch { }
        }

        private async Task DoPreviewAsync()
        {
            _btnPreview.Enabled = false;
            _btnAccept.Enabled = false;
            _cts?.Cancel();
            _cts = new CancellationTokenSource();
            var ct = _cts.Token;
            try
            {
                var ctx = new AIContext
                {
                    SheetName = _baseContext.SheetName,
                    StartRow = _baseContext.StartRow,
                    StartCol = _baseContext.StartCol,
                    Rows = (int)_numRows.Value,
                    Cols = (int)_numCols.Value,
                    Title = _baseContext.Title,
                    Prompt = _txtPrompt.Text ?? string.Empty
                };

                using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(10));
                using var linked = CancellationTokenSource.CreateLinkedTokenSource(ct, timeout.Token);
                var res = await _provider.GenerateFillAsync(ctx, linked.Token).ConfigureAwait(true);
                ResultCells = res.Cells;
                RenderPreview(res.Cells);
                _btnAccept.Enabled = true;
            }
            catch (OperationCanceledException)
            {
                // ignore
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Preview failed: {ex.Message}", "AI", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _btnPreview.Enabled = true;
            }
        }

        private void RenderPreview(string[][] cells)
        {
            _preview.Columns.Clear();
            _preview.Rows.Clear();
            int cols = cells.Length > 0 ? cells[0].Length : (int)_numCols.Value;
            for (int c = 0; c < cols; c++)
            {
                _preview.Columns.Add($"c{c}", (c + 1).ToString());
            }
            for (int r = 0; r < cells.Length; r++)
            {
                var row = new object[cols];
                for (int c = 0; c < cols; c++) row[c] = c < cells[r].Length ? cells[r][c] : string.Empty;
                _preview.Rows.Add(row);
            }
        }
    }
}
