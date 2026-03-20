using System;
using System.Globalization;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using SpreadsheetApp.Core;
using SpreadsheetApp.Core.AI;
using SpreadsheetApp.UI.AI;
using SpreadsheetApp.UI;

namespace SpreadsheetApp.UI
{
    public partial class MainForm : Form
    {
        private const string ClipboardCopyOriginFormat = "SpreadsheetApp.CopyOrigin";
        private const string ClipboardCopyContentFormat = "SpreadsheetApp.CopyContent";
        private Spreadsheet _sheet = null!;
        private readonly List<Spreadsheet> _sheets = new();
        private readonly List<string> _sheetNames = new();
        private readonly List<UndoManager> _undos = new();
        private int _activeSheetIndex = 0;
        private UndoManager _undo = new();
        private bool _suppressTabChange = false;
        private bool _suppressRecord = false;
        private FindReplaceForm? _findDlg = null;
        private string _lastFind = string.Empty;
        private string _lastReplace = string.Empty;
        private bool _lastMatchCase = false;
        private int _findPosRow = 0;
        private int _findPosCol = -1;
        private const int DefaultRows = Spreadsheet.DefaultRows;
        private const int DefaultCols = Spreadsheet.DefaultCols; // A..Z
        private IInferenceProvider _aiProvider = new MockInferenceProvider();
        private bool _aiInlineEnabled = true;
        private System.Windows.Forms.Timer _aiInlineTimer = new System.Windows.Forms.Timer();
        private System.Windows.Forms.Panel? _aiGhostPanel;
        private System.Windows.Forms.ListBox? _aiGhostList;
        private System.Windows.Forms.Label? _aiGhostStatus;
        private System.Windows.Forms.Button? _aiGhostApplyBtn;
        private System.Windows.Forms.Button? _aiGhostDismissBtn;
        private int _aiGhostStartRow = -1;
        private int _aiGhostCol = -1;
        private System.Threading.CancellationTokenSource? _aiInlineCts;
        private IChatPlanner _chatPlanner = new MockChatPlanner();
        private readonly AIActionLog _aiActionLog = new();
        private readonly System.Collections.Generic.Dictionary<string, string[]> _aiInlineCache = new();
        private SpreadsheetApp.Core.AppSettings _settings = SpreadsheetApp.Core.AppSettings.Load();
        private bool _suppressFormulaBarSync;
        // Clipboard selection outline (copy/cut visual)
        private bool _clipboardOutlineVisible;
        private bool _clipboardOutlineIsCut;
        private int _outlineTop, _outlineLeft, _outlineBottom, _outlineRight;
        // Docked Chat pane
        private System.Windows.Forms.Panel? _chatDockHost;
        private System.Windows.Forms.Splitter? _chatDockSplitter;
        private SpreadsheetApp.UI.AI.ChatAssistantPanel? _chatPane;
        private SpreadsheetApp.Core.AI.ChatSession _chatSession = new SpreadsheetApp.Core.AI.ChatSession();
        // Drag-fill handle + preview state
        private bool _dragFillActive;
        private bool _dragFillPreview;
        private System.Drawing.Rectangle _dragHandleRect;
        private int _selTop, _selLeft, _selBottom, _selRight;
        private int _previewTop, _previewLeft, _previewBottom, _previewRight;

        public MainForm()
        {
            InitializeComponent();
            // Enable double buffering on the grid to eliminate flicker
            typeof(DataGridView).InvokeMember("DoubleBuffered",
                BindingFlags.SetProperty | BindingFlags.Instance | BindingFlags.NonPublic,
                null, grid, new object[] { true });
            try { LoadRecentFiles(); } catch { }
            // Inline AI setup
            _aiInlineTimer.Interval = 200;
            _aiInlineTimer.Tick += (_, __) => { _aiInlineTimer.Stop(); _ = TriggerInlineSuggestionAsync(); };
            CreateGhostUI();
            ApplySettings(_settings);
            SyncAiMenuState();
            // Formula bar events
            _formulaBar.KeyDown += FormulaBar_KeyDown;
            _formulaBar.Enter += (_, __) => _suppressFormulaBarSync = true;
            _formulaBar.Leave += FormulaBar_Leave;
            _cellNameBox.KeyDown += CellNameBox_KeyDown;
            // Create docked chat pane (initially hidden or visible per settings)
            try { CreateDockedChatPane(); } catch { }
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.Shift | Keys.I))
            {
                AcceptInlineSuggestion();
                return true;
            }
            if (keyData == (Keys.Control | Keys.I))
            {
                OpenGenerateFill();
                return true;
            }
            if (keyData == (Keys.Control | Keys.D)) { FillDown(); return true; }
            if (keyData == (Keys.Control | Keys.R)) { FillRight(); return true; }
            if (keyData == (Keys.Control | Keys.Shift | Keys.C)) { ToggleChatPane(); return true; }
            if (keyData == (Keys.Control | Keys.Shift | Keys.F)) { _ = SmartSchemaFillAsync(); return true; }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void CreateDockedChatPane()
        {
            // host panel on the right + splitter
            _chatDockHost = new Panel
            {
                Dock = DockStyle.Right,
                Width = Math.Max(220, Math.Min(800, _settingsChatPaneWidth())),
                Visible = _settingsChatPaneVisible()
            };
            _chatDockHost.BackColor = System.Drawing.Color.FromArgb(248, 248, 248);
            _chatDockSplitter = new Splitter
            {
                Dock = DockStyle.Right,
                Width = 4,
                BackColor = System.Drawing.Color.FromArgb(234, 234, 234),
                Visible = _chatDockHost.Visible
            };
            Controls.Add(_chatDockSplitter);
            Controls.Add(_chatDockHost);

            // instantiate chat panel
            Func<AIContext> getCtx = () => BuildPlannerContext();
            void apply(AIPlan plan) => ApplyPlan(plan);
            _chatPane = new UI.AI.ChatAssistantPanel(_chatPlanner, _chatSession, getCtx, apply);
            _chatPane.Dock = DockStyle.Fill;
            _chatDockHost.Controls.Add(_chatPane);

            _chatDockSplitter.SplitterMoved += (_, __) => SaveChatPaneWidth();
            _chatDockHost.SizeChanged += (_, __) => { if (_chatDockHost.Visible) SaveChatPaneWidth(); };
        }

        private int _settingsChatPaneWidth()
        {
            try
            {
                var prop = typeof(SpreadsheetApp.Core.AppSettings).GetProperty("ChatPaneWidth");
                if (prop != null)
                {
                    var v = prop.GetValue(_settings);
                    if (v is int i && i > 0) return i;
                }
            }
            catch { }
            return 340;
        }

        private bool _settingsChatPaneVisible()
        {
            try
            {
                var prop = typeof(SpreadsheetApp.Core.AppSettings).GetProperty("ChatPaneVisible");
                if (prop != null)
                {
                    var v = prop.GetValue(_settings);
                    if (v is bool b) return b;
                }
            }
            catch { }
            return false;
        }

        private void SaveChatPaneWidth()
        {
            try
            {
                if (_chatDockHost == null) return;
                var prop = typeof(SpreadsheetApp.Core.AppSettings).GetProperty("ChatPaneWidth");
                if (prop != null) prop.SetValue(_settings, _chatDockHost.Width);
                _settings.Save();
            }
            catch { }
        }

        private void ToggleChatPane()
        {
            if (_chatDockHost == null || _chatDockSplitter == null)
            {
                try { CreateDockedChatPane(); } catch { }
            }
            if (_chatDockHost == null || _chatDockSplitter == null) return;
            bool show = !_chatDockHost.Visible;
            _chatDockHost.Visible = show;
            _chatDockSplitter.Visible = show;
            try
            {
                var prop = typeof(SpreadsheetApp.Core.AppSettings).GetProperty("ChatPaneVisible");
                if (prop != null) prop.SetValue(_settings, show);
                _settings.Save();
            }
            catch { }
            try { if (show) _chatPane?.FocusInput(); } catch { }
            try { UpdateAiMenuItemsState(); } catch { }
            // Refresh policy preview when pane becomes visible
            try
            {
                if (show && _chatPane != null)
                {
                    var ctx = BuildPlannerContext();
                    _chatPane.RefreshPolicy(ctx);
                }
            }
            catch { }
        }

        private void ShowAIActionLog()
        {
            var form = new Form
            {
                Text = "AI Action Log",
                Width = 600,
                Height = 400,
                StartPosition = FormStartPosition.CenterParent
            };
            var list = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true
            };
            list.Columns.Add("Time", 80);
            list.Columns.Add("Prompt", 200);
            list.Columns.Add("Cmds", 50);
            list.Columns.Add("Cells", 50);
            list.Columns.Add("Summary", 300);
            foreach (var e in _aiActionLog.Entries)
            {
                var item = new ListViewItem(e.Timestamp.ToString("HH:mm:ss"));
                item.SubItems.Add(e.Prompt);
                item.SubItems.Add(e.CommandCount.ToString());
                item.SubItems.Add(e.CellCount.ToString());
                item.SubItems.Add(e.Summary);
                list.Items.Add(item);
            }
            form.Controls.Add(list);
            form.ShowDialog(this);
        }

        private System.Threading.Tasks.Task ExplainCellAsync()
        {
            if (!_settings.AiEnabled || !ProviderReady()) return System.Threading.Tasks.Task.CompletedTask;
            var cell = grid.CurrentCell;
            if (cell == null) return System.Threading.Tasks.Task.CompletedTask;
            string? raw = _sheet.GetRaw(cell.RowIndex, cell.ColumnIndex);
            if (string.IsNullOrWhiteSpace(raw)) { MessageBox.Show(this, "Cell is empty.", "Explain Cell"); return System.Threading.Tasks.Task.CompletedTask; }
            var val = _sheet.GetValue(cell.RowIndex, cell.ColumnIndex);
            string addr = Core.CellAddress.ToAddress(cell.RowIndex, cell.ColumnIndex);
            string prompt = $"Explain what this cell does in plain language. Cell {addr} contains: {raw}";
            if (val.Error != null) prompt += $" (evaluates to error: {val.Error})";
            else prompt += $" (evaluates to: {val.ToDisplay()})";

            // Use planner for a simple explanation (no commands expected)
            var ctx = BuildPlannerContext();
            ctx.AllowedCommands = new string[0]; // No commands - we just want text
            try
            {
                // Ensure docked chat pane is visible and feed the prompt
                if (_chatDockHost == null || _chatDockSplitter == null) { try { CreateDockedChatPane(); } catch { } }
                if (_chatDockHost != null)
                {
                    _chatDockHost.Visible = true;
                    if (_chatDockSplitter != null) _chatDockSplitter.Visible = true;
                    _chatPane?.SetPrompt(prompt, autoPlan: true);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Error: {ex.Message}", "Explain Cell");
            }
            return System.Threading.Tasks.Task.CompletedTask;
        }

        private async System.Threading.Tasks.Task SmartSchemaFillAsync()
        {
            // If only a single cell is selected, try to auto-expand to the empty data region
            try
            {
                if (grid.SelectedCells.Count <= 1 && grid.CurrentCell != null)
                {
                    int row = grid.CurrentCell.RowIndex;
                    int col = grid.CurrentCell.ColumnIndex;
                    // Find header row (row 0 or first non-empty text row)
                    int headerRow = 0;
                    // Find the rightmost non-empty header
                    int maxCol = col;
                    for (int c = 0; c < _sheet.Columns; c++)
                    {
                        var h = _sheet.GetRaw(headerRow, c);
                        if (!string.IsNullOrWhiteSpace(h)) maxCol = c;
                    }
                    // Find the input column (leftmost non-empty column with data below headers)
                    int inputCol = 0;
                    for (int c = 0; c <= maxCol; c++)
                    {
                        if (!string.IsNullOrWhiteSpace(_sheet.GetRaw(headerRow + 1, c))) { inputCol = c; break; }
                    }
                    // Find the last row with input data
                    int lastDataRow = headerRow;
                    for (int r = headerRow + 1; r < _sheet.Rows; r++)
                    {
                        if (!string.IsNullOrWhiteSpace(_sheet.GetRaw(r, inputCol))) lastDataRow = r;
                        else break;
                    }
                    // Select the output region: from (headerRow+1, inputCol+1) to (lastDataRow, maxCol)
                    int startRow = headerRow + 1;
                    int startCol = inputCol + 1;
                    if (startCol <= maxCol && startRow <= lastDataRow)
                    {
                        grid.ClearSelection();
                        for (int r = startRow; r <= lastDataRow; r++)
                            for (int c = startCol; c <= maxCol; c++)
                                grid[c, r].Selected = true;
                        grid.CurrentCell = grid[startCol, startRow];
                    }
                }
            }
            catch { }
            await FillSelectedFromSchemaAsync();
        }

        private bool _freezeTopRow;
        private bool _freezeFirstCol;

        private void ToggleFreezeTopRow()
        {
            _freezeTopRow = !_freezeTopRow;
            if (grid.Rows.Count > 0)
                grid.Rows[0].Frozen = _freezeTopRow;
        }

        private void ToggleFreezeFirstColumn()
        {
            _freezeFirstCol = !_freezeFirstCol;
            if (grid.Columns.Count > 0)
                grid.Columns[0].Frozen = _freezeFirstCol;
        }

        private void OpenDocsViewer()
        {
            try
            {
                using var dlg = new DocsViewerForm();
                dlg.ShowDialog(this);
            }
            catch (Exception ex)
            {
                try { MessageBox.Show(this, $"Failed to open Docs Viewer: {ex.Message}", "Docs Viewer", MessageBoxButtons.OK, MessageBoxIcon.Error); } catch { }
            }
        }

        private void ExportDocsJson()
        {
            try
            {
                var idx = DocsIndexer.Build(null);
                string path = DocsIndexer.WriteJson(idx, null);
                try { MessageBox.Show(this, $"Exported to:\n{path}", "Docs Export", MessageBoxButtons.OK, MessageBoxIcon.Information); } catch { }
            }
            catch (Exception ex)
            {
                try { MessageBox.Show(this, $"Failed to export: {ex.Message}", "Docs Export", MessageBoxButtons.OK, MessageBoxIcon.Error); } catch { }
            }
        }

        private void InitializeSheet(Spreadsheet? newSheet = null)
        {
            if (_sheets.Count == 0)
            {
                _sheets.Add(newSheet ?? new Spreadsheet(DefaultRows, DefaultCols));
                _sheetNames.Add("Sheet1");
                _undos.Add(new UndoManager());
                _activeSheetIndex = 0;
            }
            else if (newSheet != null)
            {
                _sheets[_activeSheetIndex] = newSheet;
                _undos[_activeSheetIndex].Clear();
            }
            _sheet = _sheets[_activeSheetIndex];
            _undo = _undos[_activeSheetIndex];
            // Wire cross-sheet resolver for =Sheet2!A1 formulas
            WireCrossSheetResolver(_sheet);

            grid.Columns.Clear();
            grid.Rows.Clear();

            for (int c = 0; c < _sheet.Columns; c++)
            {
                var col = new DataGridViewTextBoxColumn();
                col.Name = CellAddress.ColumnIndexToName(c);
                col.HeaderText = col.Name;
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
                col.Width = 80;
                grid.Columns.Add(col);
            }

            grid.Rows.Add(_sheet.Rows);
            for (int r = 0; r < _sheet.Rows; r++)
            {
                grid.Rows[r].HeaderCell.Value = (r + 1).ToString(CultureInfo.InvariantCulture);
            }

            RefreshGridValues();
            UpdateStatus();
            RefreshTabs();
            try { UpdateAiMenuItemsState(); } catch { }
        }

        private void WireCrossSheetResolver(Spreadsheet sheet)
        {
            sheet.CrossSheetResolver = (sheetName, row, col) =>
            {
                for (int i = 0; i < _sheets.Count; i++)
                {
                    if (string.Equals(_sheetNames[i], sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        var target = _sheets[i];
                        if (row < 0 || row >= target.Rows || col < 0 || col >= target.Columns)
                            return EvaluationResult.FromError("REF");
                        return target.GetValue(row, col);
                    }
                }
                return EvaluationResult.FromError($"Sheet '{sheetName}' not found");
            };
        }

        private void Grid_CellBeginEdit(object? sender, DataGridViewCellCancelEventArgs e)
        {
            var raw = _sheet.GetRaw(e.RowIndex, e.ColumnIndex);
            grid[e.ColumnIndex, e.RowIndex].Value = raw ?? string.Empty;
            UpdateStatus();
        }

        private void Grid_CellEndEdit(object? sender, DataGridViewCellEventArgs e)
        {
            var cell = grid[e.ColumnIndex, e.RowIndex];
            var raw = cell.Value?.ToString() ?? string.Empty;
            var oldRaw = _sheet.GetRaw(e.RowIndex, e.ColumnIndex);
            _sheet.SetRaw(e.RowIndex, e.ColumnIndex, raw);
            if (!_suppressRecord)
            {
                _undo.RecordSet(e.RowIndex, e.ColumnIndex, oldRaw, raw);
            }
            try
            {
                // Incremental-on-edit: recalc and repaint affected cells with fallback inside RefreshDirtyOrFull
                var affected = _sheet.RecalculateDirty(e.RowIndex, e.ColumnIndex);
                RefreshDirtyOrFull(affected);
                UpdateStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Error during calculation: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            // Schedule inline suggestions after edit
            if (_aiInlineEnabled) ScheduleInlineSuggestion();
            try { UpdateAiMenuItemsState(); } catch { }
        }

        private void Undo()
        {
            if (_undo.TryUndoComposite(out var editsComp, out var sheetAddComp))
            {
                // Undo bulk edits first (set to old values), then undo sheet adds
                if (editsComp.Count > 0)
                {
                    _suppressRecord = true;
                    var affected = new HashSet<(int r, int c)>();
                    try
                    {
                        foreach (var e in editsComp)
                        {
                            _sheet.SetRaw(e.Row, e.Col, e.OldRaw);
                            foreach (var ac in _sheet.RecalculateDirty(e.Row, e.Col)) affected.Add(ac);
                        }
                    }
                    finally { _suppressRecord = false; }
                    try { RefreshDirtyOrFull(affected); } catch { }
                }
                if (sheetAddComp.HasValue)
                {
                    RemoveSheetAt(sheetAddComp.Value.index);
                }
                // Removing a sheet triggers full re-init of grid
            }
            else if (_undo.TryUndoSheetAdd(out int sheetIndex, out string sheetName))
            {
                RemoveSheetAt(sheetIndex);
            }
            else if (_undo.TryUndoBulk(out var edits))
            {
                _suppressRecord = true;
                var affected = new HashSet<(int r, int c)>();
                try
                {
                    foreach (var e in edits)
                    {
                        _sheet.SetRaw(e.row, e.col, e.raw);
                        foreach (var ac in _sheet.RecalculateDirty(e.row, e.col)) affected.Add(ac);
                    }
                }
                finally { _suppressRecord = false; }
                try { RefreshDirtyOrFull(affected); } catch { }
                if (edits.Count > 0)
                {
                    var (rr, cc, _) = edits[0];
                    if (rr >= 0 && rr < grid.RowCount && cc >= 0 && cc < grid.ColumnCount)
                        grid.CurrentCell = grid[cc, rr];
                }
            }
            else if (_undo.TryUndo(out int r, out int c, out string? raw))
            {
                _suppressRecord = true;
                _sheet.SetRaw(r, c, raw);
                try { var aff = _sheet.RecalculateDirty(r, c); RefreshDirtyOrFull(aff); } finally { _suppressRecord = false; }
                if (r >= 0 && r < grid.RowCount && c >= 0 && c < grid.ColumnCount)
                {
                    grid.CurrentCell = grid[c, r];
                }
            }
            try { UpdateAiMenuItemsState(); } catch { }
        }

        private void Redo()
        {
            if (_undo.TryRedoComposite(out var editsComp, out var sheetAddComp))
            {
                // Redo sheet adds first, then apply edits to their new values
                if (sheetAddComp.HasValue)
                {
                    AddSheetAt(sheetAddComp.Value.index, sheetAddComp.Value.name);
                }
                if (editsComp.Count > 0)
                {
                    _suppressRecord = true;
                    var affected = new HashSet<(int r, int c)>();
                    try
                    {
                        foreach (var e in editsComp)
                        {
                            _sheet.SetRaw(e.Row, e.Col, e.NewRaw);
                            foreach (var ac in _sheet.RecalculateDirty(e.Row, e.Col)) affected.Add(ac);
                        }
                    }
                    finally { _suppressRecord = false; }
                    try { RefreshDirtyOrFull(affected); } catch { }
                }
            }
            else if (_undo.TryRedoSheetAdd(out int sheetIndex, out string sheetName))
            {
                AddSheetAt(sheetIndex, sheetName);
            }
            else if (_undo.TryRedoBulk(out var edits))
            {
                _suppressRecord = true;
                var affected = new HashSet<(int r, int c)>();
                try
                {
                    foreach (var e in edits)
                    {
                        _sheet.SetRaw(e.row, e.col, e.raw);
                        foreach (var ac in _sheet.RecalculateDirty(e.row, e.col)) affected.Add(ac);
                    }
                }
                finally { _suppressRecord = false; }
                try { RefreshDirtyOrFull(affected); } catch { }
                if (edits.Count > 0)
                {
                    var (rr, cc, _) = edits[0];
                    if (rr >= 0 && rr < grid.RowCount && cc >= 0 && cc < grid.ColumnCount)
                        grid.CurrentCell = grid[cc, rr];
                }
            }
            else if (_undo.TryRedo(out int r, out int c, out string? raw))
            {
                _suppressRecord = true;
                _sheet.SetRaw(r, c, raw);
                try { var aff = _sheet.RecalculateDirty(r, c); RefreshDirtyOrFull(aff); } finally { _suppressRecord = false; }
                if (r >= 0 && r < grid.RowCount && c >= 0 && c < grid.ColumnCount)
                {
                    grid.CurrentCell = grid[c, r];
                }
            }
            try { UpdateAiMenuItemsState(); } catch { }
        }

        private void RefreshGridValues()
        {
            _sheet.Recalculate();
            RefreshGridDisplays();
        }

        private void RefreshDirtyOrFull(System.Collections.Generic.IReadOnlyCollection<(int r, int c)> affected)
        {
            try
            {
                int total = _sheet.Rows * _sheet.Columns;
                if (affected == null || affected.Count == 0 || affected.Count > Math.Max(10, total / 20))
                {
                    RefreshGridValues();
                    grid.Invalidate();
                    grid.Refresh();
                    UpdateStatus();
                    return;
                }
                grid.SuspendLayout();
                try
                {
                    foreach (var (r, c) in affected)
                    {
                        if (r < 0 || r >= _sheet.Rows || c < 0 || c >= _sheet.Columns) continue;
                        var cell = grid[c, r];
                        cell.Value = GetDisplayWithFormat(r, c);
                        ApplyCellFormat(cell, r, c);
                        grid.InvalidateCell(c, r);
                    }
                }
                finally
                {
                    grid.ResumeLayout();
                    // Force immediate repaint so the UI reflects the latest step during Test Runner automation
                    try { grid.Refresh(); } catch { }
                    UpdateStatus();
                }
            }
            catch
            {
                try { RefreshGridValues(); grid.Invalidate(); grid.Refresh(); UpdateStatus(); } catch { }
            }
        }

        private void RefreshGridDisplays()
        {
            for (int r = 0; r < _sheet.Rows; r++)
            {
                for (int c = 0; c < _sheet.Columns; c++)
                {
                    var disp = GetDisplayWithFormat(r, c);
                    var cell = grid[c, r];
                    cell.Value = disp;
                    ApplyCellFormat(cell, r, c);
                }
            }
        }

        private string GetDisplayWithFormat(int r, int c)
        {
            var fmt = _sheet.GetFormat(r, c);
            if (fmt != null && !string.IsNullOrEmpty(fmt.NumberFormat))
            {
                try
                {
                    var val = _sheet.GetValue(r, c);
                    if (val.Error == null && val.Number is double d)
                    {
                        string f = fmt.NumberFormat!;
                        switch (f)
                        {
                            case "0":
                                return d.ToString("0", CultureInfo.InvariantCulture);
                            case "0.00":
                                return d.ToString("0.00", CultureInfo.InvariantCulture);
                            case "#,##0":
                                return d.ToString("#,##0", CultureInfo.InvariantCulture);
                            case "#,##0.00":
                                return d.ToString("#,##0.00", CultureInfo.InvariantCulture);
                            case "0%":
                                return d.ToString("0%", CultureInfo.InvariantCulture);
                            case "0.00%":
                                return d.ToString("0.00%", CultureInfo.InvariantCulture);
                            case "$#,##0":
                            {
                                string core = d.ToString("#,##0", CultureInfo.InvariantCulture);
                                return d < 0 ? "-$" + core.TrimStart('-') : "$" + core;
                            }
                            case "$#,##0.00":
                            {
                                string core = d.ToString("#,##0.00", CultureInfo.InvariantCulture);
                                return d < 0 ? "-$" + core.TrimStart('-') : "$" + core;
                            }
                        }
                    }
                }
                catch { }
            }
            return _sheet.GetDisplay(r, c);
        }

        private void ApplyCellFormat(DataGridViewCell cell, int r, int c)
        {
            var fmt = _sheet.GetFormat(r, c);
            var style = cell.Style;
            // Font bold
            if (fmt != null && fmt.Bold)
                style.Font = new System.Drawing.Font(grid.Font, System.Drawing.FontStyle.Bold);
            else
                style.Font = grid.Font;

            // Foreground / background
            if (fmt != null && fmt.ForeColorArgb.HasValue)
                style.ForeColor = System.Drawing.Color.FromArgb(fmt.ForeColorArgb.Value);
            else
                style.ForeColor = System.Drawing.Color.Empty;

            if (fmt != null && fmt.BackColorArgb.HasValue)
                style.BackColor = System.Drawing.Color.FromArgb(fmt.BackColorArgb.Value);
            else
                style.BackColor = System.Drawing.Color.Empty;

            // Alignment
            if (fmt != null)
            {
                style.Alignment = fmt.HAlign switch
                {
                    CellHAlign.Center => DataGridViewContentAlignment.MiddleCenter,
                    CellHAlign.Right => DataGridViewContentAlignment.MiddleRight,
                    _ => DataGridViewContentAlignment.MiddleLeft
                };
            }
            else
            {
                style.Alignment = DataGridViewContentAlignment.MiddleLeft;
            }
        }

        private void Grid_SelectionChanged(object? sender, EventArgs e)
        {
            UpdateStatus();
            HideGhostSuggestions();
            if (_aiInlineEnabled) ScheduleInlineSuggestion();
            // Keep the docked chat pane policy in sync with current selection
            try
            {
                if (_chatDockHost != null && _chatDockHost.Visible && _chatPane != null)
                {
                    var ctx = BuildPlannerContext();
                    _chatPane.RefreshPolicy(ctx);
                }
            }
            catch { }
            // Redraw to update drag handle
            try { grid.Invalidate(); } catch { }
        }

        private void UpdateStatus()
        {
            if (grid.CurrentCell == null)
            {
                statusCell.Text = "Cell: -";
                statusRaw.Text = "Raw: ";
                statusValue.Text = "Value: ";
                if (!_suppressFormulaBarSync)
                {
                    _cellNameBox.Text = string.Empty;
                    _formulaBar.Text = string.Empty;
                }
                return;
            }
            int r = grid.CurrentCell.RowIndex;
            int c = grid.CurrentCell.ColumnIndex;
            string addr = CellAddress.ToAddress(r, c);
            string raw = _sheet.GetRaw(r, c) ?? string.Empty;
            string val;
            try { val = GetDisplayWithFormat(r, c); }
            catch { val = string.Empty; }
            statusCell.Text = $"Cell: {addr}";
            statusRaw.Text = $"Raw: {raw}";
            statusValue.Text = $"Value: {val}";
            if (!_suppressFormulaBarSync)
            {
                _cellNameBox.Text = addr;
                _formulaBar.Text = raw;
            }
        }

        // --- Formula Bar ---
        private void FormulaBar_KeyDown(object? sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                CommitFormulaBarEdit();
            }
            else if (e.KeyCode == Keys.Escape)
            {
                e.SuppressKeyPress = true;
                _suppressFormulaBarSync = false;
                UpdateStatus();
                grid.Focus();
            }
            else if (e.KeyCode == Keys.Tab)
            {
                e.SuppressKeyPress = true;
                CommitFormulaBarEdit();
                // Move to next cell
                if (grid.CurrentCell != null)
                {
                    int r = grid.CurrentCell.RowIndex;
                    int c = grid.CurrentCell.ColumnIndex + 1;
                    if (c < grid.ColumnCount)
                    {
                        grid.CurrentCell = grid[c, r];
                    }
                }
            }
        }

        private void CommitFormulaBarEdit()
        {
            if (grid.CurrentCell == null) { grid.Focus(); return; }
            int r = grid.CurrentCell.RowIndex;
            int c = grid.CurrentCell.ColumnIndex;
            string newVal = _formulaBar.Text ?? string.Empty;
            string oldVal = _sheet.GetRaw(r, c) ?? string.Empty;
            if (newVal != oldVal)
            {
                _sheet.SetRaw(r, c, newVal);
                _undo.RecordSet(r, c, oldVal, newVal);
                _sheet.RecalculateDirty(r, c);
                RefreshGridValues();
            }
            _suppressFormulaBarSync = false;
            UpdateStatus();
            grid.Focus();
        }

        private void FormulaBar_Leave(object? sender, EventArgs e)
        {
            _suppressFormulaBarSync = false;
            UpdateStatus();
        }

        private void CellNameBox_KeyDown(object? sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                string addr = _cellNameBox.Text?.Trim() ?? string.Empty;
                if (CellAddress.TryParse(addr, out int row, out int col))
                {
                    if (row < grid.RowCount && col < grid.ColumnCount)
                    {
                        grid.CurrentCell = grid[col, row];
                    }
                }
                grid.Focus();
            }
            else if (e.KeyCode == Keys.Escape)
            {
                e.SuppressKeyPress = true;
                _suppressFormulaBarSync = false;
                UpdateStatus();
                grid.Focus();
            }
        }

        // --- Copy / Paste / Cut ---
        private void CopyCell()
        {
            if (grid.CurrentCell == null) return;
            // Compute rectangular selection bounds (or current cell if single)
            TryGetSelectionBounds(out int top, out int left, out int bottom, out int right);
            var sb = new System.Text.StringBuilder();
            var vals2d = new System.Collections.Generic.List<string[]>();
            for (int r = top; r <= bottom; r++)
            {
                var rowVals = new string[right - left + 1];
                for (int c = left; c <= right; c++)
                {
                    if (c > left) sb.Append('\t');
                    var raw = _sheet.GetRaw(r, c) ?? string.Empty;
                    // Escape tabs/newlines minimally by quoting if needed
                    if (raw.Contains('\t') || raw.Contains('\n') || raw.Contains('\r'))
                    {
                        string q = '"' + raw.Replace("\"", "\"\"") + '"';
                        sb.Append(q);
                    }
                    else sb.Append(raw);
                    rowVals[c - left] = raw;
                }
                if (r < bottom) sb.AppendLine();
                vals2d.Add(rowVals);
            }
            try
            {
                var dobj = new DataObject();
                dobj.SetText(sb.ToString());
                dobj.SetData(ClipboardCopyOriginFormat, $"{top},{left}");
                // Include structured raw values payload for reliable internal pastes
                try
                {
                    var payload = new { rows = bottom - top + 1, cols = right - left + 1, values = vals2d.ToArray() };
                    string json = System.Text.Json.JsonSerializer.Serialize(payload);
                    dobj.SetData(ClipboardCopyContentFormat, json);
                }
                catch { }
                Clipboard.SetDataObject(dobj, true);
            }
            catch { /* ignore */ }
            // Show copy outline
            ShowClipboardOutline(top, left, bottom, right, isCut: false);
        }

        private void PasteCell()
        {
            if (grid.CurrentCell == null) return;
            try { if (grid.IsCurrentCellInEditMode) grid.EndEdit(); } catch { }
            string text = string.Empty;
            try { if (Clipboard.ContainsText()) text = Clipboard.GetText(); } catch { }
            if (string.IsNullOrEmpty(text)) return;

            // Parse TSV/CSV-like text into rows x cols
            static string Unquote(string s)
            {
                if (s.Length >= 2 && s[0] == '"' && s[^1] == '"')
                {
                    var inner = s.Substring(1, s.Length - 2);
                    return inner.Replace("\"\"", "\"");
                }
                return s;
            }

            var lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            // Drop a trailing empty line if present (common when copying)
            if (lines.Length > 0 && lines[^1].Length == 0)
            {
                Array.Resize(ref lines, lines.Length - 1);
            }
            string[][] table;
            // Prefer internal structured payload when present (preserves raw formulas accurately)
            bool usedStructured = false;
            try
            {
                var dobj = Clipboard.GetDataObject();
                if (dobj != null && dobj.GetDataPresent(ClipboardCopyContentFormat))
                {
                    var json = dobj.GetData(ClipboardCopyContentFormat)?.ToString();
                    if (!string.IsNullOrWhiteSpace(json))
                    {
                        var doc = System.Text.Json.JsonDocument.Parse(json);
                        if (doc.RootElement.TryGetProperty("values", out var vals) && vals.ValueKind == System.Text.Json.JsonValueKind.Array)
                        {
                            var rowsList = new System.Collections.Generic.List<string[]>();
                            foreach (var rowEl in vals.EnumerateArray())
                            {
                                var colsList = new System.Collections.Generic.List<string>();
                                if (rowEl.ValueKind == System.Text.Json.JsonValueKind.Array)
                                {
                                    foreach (var c in rowEl.EnumerateArray()) colsList.Add(c.GetString() ?? string.Empty);
                                }
                                rowsList.Add(colsList.ToArray());
                            }
                            table = rowsList.ToArray();
                            usedStructured = true;
                        }
                        else table = Array.Empty<string[]>();
                    }
                    else table = Array.Empty<string[]>();
                }
                else
                {
                    table = Array.Empty<string[]>();
                }
            }
            catch { table = Array.Empty<string[]>(); }
            if (!usedStructured)
            {
                if (lines.Length <= 1 && !text.Contains('\t'))
                {
                    // Single cell paste
                    table = new[] { new[] { text } };
                }
                else
                {
                    var rowsList = new List<string[]>();
                    foreach (var line in lines)
                    {
                        var parts = line.Split('\t');
                        for (int i = 0; i < parts.Length; i++) parts[i] = Unquote(parts[i]);
                        rowsList.Add(parts);
                    }
                    table = rowsList.ToArray();
                }
            }

            // Determine paste start at current cell (anchor). This avoids surprises from lingering multi-selections.
            int startR = grid.CurrentCell.RowIndex;
            int startC = grid.CurrentCell.ColumnIndex;

            // Compute rewrite delta if clipboard originated from this app
            bool canRewrite = false; int dr = 0, dc = 0;
            try
            {
                if (Clipboard.GetDataObject() is IDataObject dobj && dobj.GetDataPresent(ClipboardCopyOriginFormat))
                {
                    var s = dobj.GetData(ClipboardCopyOriginFormat)?.ToString();
                    if (!string.IsNullOrWhiteSpace(s))
                    {
                        var parts = s.Split(',');
                        if (parts.Length == 2 && int.TryParse(parts[0], out int srcR) && int.TryParse(parts[1], out int srcC))
                        {
                            dr = startR - srcR; dc = startC - srcC; canRewrite = true;
                        }
                    }
                }
            }
            catch { }

            var edits = new List<(int row, int col, string? oldRaw, string? newRaw)>();
            var affected = new HashSet<(int r, int c)>();
            _suppressRecord = true;
            try
            {
                for (int r = 0; r < table.Length; r++)
                {
                    var row = table[r];
                    int cols = row.Length;
                    for (int c = 0; c < cols; c++)
                    {
                        int rr = startR + r;
                        int cc = startC + c;
                        if (rr < 0 || rr >= _sheet.Rows || cc < 0 || cc >= _sheet.Columns) continue;
                        string? oldRaw = _sheet.GetRaw(rr, cc);
                        string newRaw = row[c] ?? string.Empty;
                        if (canRewrite && !string.IsNullOrEmpty(newRaw))
                        {
                            if (newRaw.StartsWith("=", StringComparison.Ordinal))
                            {
                                try { newRaw = RewriteFormulaForPaste(newRaw, dr, dc); } catch { }
                            }
                            else if (LooksLikeCellRefOrRange(newRaw))
                            {
                                try { newRaw = RewriteFormulaForPaste("=" + newRaw, dr, dc); } catch { }
                            }
                        }
                        if (oldRaw == newRaw) continue;
                        _sheet.SetRaw(rr, cc, newRaw);
                        edits.Add((rr, cc, oldRaw, newRaw));
                        foreach (var ac in _sheet.RecalculateDirty(rr, cc)) affected.Add(ac);
                    }
                }
            }
            finally { _suppressRecord = false; }
            if (edits.Count > 0)
            {
                _undo.RecordBulk(edits);
                try { RefreshDirtyOrFull(affected); } catch { }
            }
            try { UpdateAiMenuItemsState(); } catch { }
            // Clear copy/cut outline after paste
            ClearClipboardOutline();
        }

        private static bool TryParseCellRefWithAnchors(string s, ref int i, out int row, out int col, out bool absCol, out bool absRow, out int tokenStart, out int tokenEnd)
        {
            row = col = 0; absCol = absRow = false; tokenStart = i; tokenEnd = i;
            int pos = i;
            int save = i;
            if (pos < s.Length && s[pos] == '$') { absCol = true; pos++; }
            int lettersStart = pos;
            while (pos < s.Length && char.IsLetter(s[pos])) pos++;
            if (pos == lettersStart) { i = save; return false; }
            if (pos < s.Length && s[pos] == '$') { absRow = true; pos++; }
            int digitsStart = pos;
            while (pos < s.Length && char.IsDigit(s[pos])) pos++;
            if (pos == digitsStart) { i = save; return false; }
            string letters = s.Substring(lettersStart, pos - lettersStart).Replace("$", string.Empty);
            string digits = s.Substring(digitsStart, pos - digitsStart);
            string addr = (letters + digits).ToUpperInvariant();
            if (!Core.CellAddress.TryParse(addr, out row, out col)) { i = save; return false; }
            tokenStart = i; tokenEnd = pos; i = pos; return true;
        }

        private static string BuildRefWithAnchors(int row, int col, bool absCol, bool absRow)
        {
            string colName = Core.CellAddress.ColumnIndexToName(Math.Max(0, col));
            string rowText = (Math.Max(0, row) + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
            return (absCol ? "$" : string.Empty) + colName + (absRow ? "$" : string.Empty) + rowText;
        }

        private string RewriteFormulaForPaste(string formula, int dr, int dc)
        {
            // Rewrite references using a regex outside string literals. Handles single refs and ranges; preserves $ anchors.
            string expr = formula ?? string.Empty;
            var result = new System.Text.StringBuilder(expr.Length + 16);
            int i = 0;
            while (i < expr.Length)
            {
                if (expr[i] == '"')
                {
                    // Copy string literal verbatim, respecting doubled quotes
                    int start = i; i++;
                    while (i < expr.Length)
                    {
                        if (expr[i] == '"')
                        {
                            if (i + 1 < expr.Length && expr[i + 1] == '"') { i += 2; continue; }
                            i++; break;
                        }
                        i++;
                    }
                    result.Append(expr, start, i - start);
                    continue;
                }
                // Non-literal segment
                int segStart = i;
                while (i < expr.Length && expr[i] != '"') i++;
                string segment = expr.Substring(segStart, i - segStart);
                segment = RewriteRefsInSegment(segment, dr, dc);
                result.Append(segment);
            }
            return result.ToString();
        }

        private static string RewriteRefsInSegment(string segment, int dr, int dc)
        {
            if (string.IsNullOrEmpty(segment)) return segment;
            // Combined ref or range: ($?Letters)( $?Digits )( : ($?Letters)( $?Digits ) )?
            var rx = new System.Text.RegularExpressions.Regex(@"(\$?[A-Za-z]+)(\$?\d+)(\s*:\s*(\$?[A-Za-z]+)(\$?\d+))?", System.Text.RegularExpressions.RegexOptions.Compiled);
            string Eval(System.Text.RegularExpressions.Match m)
            {
                string Render(string colTok, string rowTok)
                {
                    bool absC = colTok.StartsWith("$");
                    bool absR = rowTok.StartsWith("$");
                    string colName = absC ? colTok.Substring(1) : colTok;
                    string rowDigits = absR ? rowTok.Substring(1) : rowTok;
                    if (!int.TryParse(rowDigits, out int row1)) return m.Value; // fallback
                    int r = Math.Max(0, row1 - 1);
                    int c;
                    try { c = Core.CellAddress.ColumnNameToIndex(colName); }
                    catch { return m.Value; }
                    int nr = absR ? r : r + dr;
                    int nc = absC ? c : c + dc;
                    return BuildRefWithAnchors(nr, nc, absC, absR);
                }
                if (m.Groups[3].Success)
                {
                    // Range
                    string left = Render(m.Groups[1].Value, m.Groups[2].Value);
                    string right = Render(m.Groups[4].Value, m.Groups[5].Value);
                    return left + ":" + right;
                }
                // Single ref
                return Render(m.Groups[1].Value, m.Groups[2].Value);
            }
            return rx.Replace(segment, new System.Text.RegularExpressions.MatchEvaluator(Eval));
        }

        private static bool LooksLikeCellRefOrRange(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            int i = 0;
            int r, c; bool ac, ar; int ts, te;
            if (!TryParseCellRefWithAnchors(s, ref i, out r, out c, out ac, out ar, out ts, out te)) return false;
            // Optional whitespace
            while (i < s.Length && char.IsWhiteSpace(s[i])) i++;
            if (i >= s.Length) return true;
            if (s[i] != ':') return i == s.Length;
            i++;
            return TryParseCellRefWithAnchors(s, ref i, out r, out c, out ac, out ar, out ts, out te) && i == s.Length;
        }

        private void CutCell()
        {
            if (grid.CurrentCell == null) return;
            // Copy current selection to clipboard
            // Capture selection bounds before contents are cleared
            TryGetSelectionBounds(out int top, out int left, out int bottom, out int right);
            CopyCell();
            // Then clear selected cells (with confirmation for formulas)
            ClearSelectedCells();
            try { UpdateAiMenuItemsState(); } catch { }
            // Show cut outline (dashed)
            ShowClipboardOutline(top, left, bottom, right, isCut: true);
        }

        // Keyboard navigation: Enter down, Tab right when not editing
        private void Grid_KeyDown(object? sender, KeyEventArgs e)
        {
            if (grid.CurrentCell == null) return;
            if (grid.IsCurrentCellInEditMode) return; // let edit handle keys
            // Accept inline ghost suggestion with Enter/Tab before default navigation
            if (_aiInlineEnabled && _aiGhostPanel != null && _aiGhostPanel.Visible && (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab))
            {
                e.Handled = true; e.SuppressKeyPress = true; AcceptInlineSuggestion(); return;
            }
            if (_aiInlineEnabled && _aiGhostPanel != null && _aiGhostPanel.Visible && e.KeyCode == Keys.Escape)
            {
                e.Handled = true; e.SuppressKeyPress = true; HideGhostSuggestions(); return;
            }
            if (e.KeyCode == Keys.Escape)
            {
                e.Handled = true; e.SuppressKeyPress = true;
                _dragFillActive = false; _dragFillPreview = false; try { grid.Invalidate(); } catch { }
                ClearClipboardOutline();
                return;
            }
            // Clear contents of all selected cells with Delete/Backspace
            if (e.KeyCode == Keys.Delete || e.KeyCode == Keys.Back)
            {
                ClearSelectedCells();
                e.Handled = true;
                e.SuppressKeyPress = true;
                return;
            }
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                int r = grid.CurrentCell.RowIndex;
                int c = grid.CurrentCell.ColumnIndex;
                int nr = Math.Min(r + 1, grid.RowCount - 1);
                grid.CurrentCell = grid[c, nr];
                UpdateStatus();
            }
            else if (e.KeyCode == Keys.Tab)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                int r = grid.CurrentCell.RowIndex;
                int c = grid.CurrentCell.ColumnIndex;
                int nc = c + 1;
                int nr = r;
                if (nc >= grid.ColumnCount)
                {
                    nc = 0;
                    nr = Math.Min(r + 1, grid.RowCount - 1);
                }
                grid.CurrentCell = grid[nc, nr];
                UpdateStatus();
            }
            else if (e.Control && e.KeyCode == Keys.C)
            {
                e.Handled = true; e.SuppressKeyPress = true; CopyCell();
            }
            else if (e.Control && e.KeyCode == Keys.V)
            {
                e.Handled = true; e.SuppressKeyPress = true; PasteCell();
            }
            else if (e.Control && e.KeyCode == Keys.X)
            {
                e.Handled = true; e.SuppressKeyPress = true; CutCell();
            }
            else if (e.Control && e.KeyCode == Keys.F)
            {
                e.Handled = true; e.SuppressKeyPress = true; OpenFindDialog(false);
            }
            else if (e.Control && e.KeyCode == Keys.H)
            {
                e.Handled = true; e.SuppressKeyPress = true; OpenFindDialog(true);
            }
            else if (e.Control && e.Shift && e.KeyCode == Keys.I)
            {
                e.Handled = true; e.SuppressKeyPress = true; AcceptInlineSuggestion();
            }
            else if (e.Control && !e.Shift && e.KeyCode == Keys.I)
            {
                e.Handled = true; e.SuppressKeyPress = true; OpenGenerateFill();
            }
            else if (_aiInlineEnabled && _aiGhostPanel != null && _aiGhostPanel.Visible && (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab))
            {
                e.Handled = true; e.SuppressKeyPress = true; AcceptInlineSuggestion();
            }
            // Ignore pure modifier keys so the ghost list stays visible while composing hotkeys
            else if (e.KeyCode == Keys.ControlKey || e.KeyCode == Keys.ShiftKey || e.KeyCode == Keys.Menu)
            {
                return;
            }
            else
            {
                // Any other key likely invalidates suggestions
                HideGhostSuggestions();
                CancelInlineRequest();
            }
        }

        private void ClearSelectedCells()
        {
            try
            {
                var cells = grid.SelectedCells;
                if (cells == null || cells.Count == 0) return;
                // Optional safety: confirm if formulas would be cleared
                int formulaCount = 0;
                foreach (DataGridViewCell cell in cells)
                {
                    if (cell == null) continue;
                    var raw = _sheet.GetRaw(cell.RowIndex, cell.ColumnIndex) ?? string.Empty;
                    if (!string.IsNullOrEmpty(raw) && raw.StartsWith("=", StringComparison.Ordinal)) formulaCount++;
                }
                if (formulaCount > 0)
                {
                    var resp = MessageBox.Show(this, $"Clear contents of {cells.Count} cells, including {formulaCount} formula(s)?", "Clear Contents", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    if (resp != DialogResult.OK) return;
                }
                var seen = new HashSet<(int r, int c)>();
                var edits = new List<(int row, int col, string? oldRaw, string? newRaw)>();
                var affected = new HashSet<(int r, int c)>();
                foreach (DataGridViewCell cell in cells)
                {
                    if (cell == null) continue;
                    int r = cell.RowIndex, c = cell.ColumnIndex;
                    if (!seen.Add((r, c))) continue;
                    string? oldRaw = _sheet.GetRaw(r, c);
                    if (!string.IsNullOrEmpty(oldRaw))
                    {
                        _sheet.SetRaw(r, c, string.Empty);
                        edits.Add((r, c, oldRaw, string.Empty));
                        foreach (var ac in _sheet.RecalculateDirty(r, c)) affected.Add(ac);
                    }
                }
                if (edits.Count > 0)
                {
                    _undo.RecordBulk(edits);
                    RefreshDirtyOrFull(affected);
                    try { UpdateAiMenuItemsState(); } catch { }
                }
            }
            catch { }
        }

        // --- Fill Down / Fill Right ---
        private void FillDown()
        {
            if (grid.CurrentCell == null) return;
            if (grid.IsCurrentCellInEditMode) return;
            TryGetSelectionBounds(out int top, out int left, out int bottom, out int right);
            if (bottom <= top) return; // nothing to fill down
            var edits = new List<(int row, int col, string? oldRaw, string? newRaw)>();
            var affected = new HashSet<(int r, int c)>();
            _suppressRecord = true;
            try
            {
                for (int c = left; c <= right; c++)
                {
                    string src = _sheet.GetRaw(top, c) ?? string.Empty;
                    for (int r = top + 1; r <= bottom; r++)
                    {
                        string newRaw = src;
                        int dr = r - top; int dc = 0;
                        if (!string.IsNullOrEmpty(newRaw))
                        {
                            if (newRaw.StartsWith("=", StringComparison.Ordinal))
                            {
                                try { newRaw = RewriteFormulaForPaste(newRaw, dr, dc); } catch { }
                            }
                            else if (LooksLikeCellRefOrRange(newRaw))
                            {
                                try { newRaw = RewriteFormulaForPaste("=" + newRaw, dr, dc); } catch { }
                            }
                        }
                        string? oldRaw = _sheet.GetRaw(r, c);
                        if (oldRaw == newRaw) continue;
                        _sheet.SetRaw(r, c, newRaw);
                        edits.Add((r, c, oldRaw, newRaw));
                        foreach (var ac in _sheet.RecalculateDirty(r, c)) affected.Add(ac);
                    }
                }
            }
            finally { _suppressRecord = false; }
            if (edits.Count > 0)
            {
                _undo.RecordBulk(edits);
                try { RefreshDirtyOrFull(affected); } catch { }
                try { UpdateAiMenuItemsState(); } catch { }
            }
        }

        private void FillRight()
        {
            if (grid.CurrentCell == null) return;
            if (grid.IsCurrentCellInEditMode) return;
            TryGetSelectionBounds(out int top, out int left, out int bottom, out int right);
            if (right <= left) return; // nothing to fill right
            var edits = new List<(int row, int col, string? oldRaw, string? newRaw)>();
            var affected = new HashSet<(int r, int c)>();
            _suppressRecord = true;
            try
            {
                for (int r = top; r <= bottom; r++)
                {
                    string src = _sheet.GetRaw(r, left) ?? string.Empty;
                    for (int c = left + 1; c <= right; c++)
                    {
                        string newRaw = src;
                        int dr = 0; int dc = c - left;
                        if (!string.IsNullOrEmpty(newRaw))
                        {
                            if (newRaw.StartsWith("=", StringComparison.Ordinal))
                            {
                                try { newRaw = RewriteFormulaForPaste(newRaw, dr, dc); } catch { }
                            }
                            else if (LooksLikeCellRefOrRange(newRaw))
                            {
                                try { newRaw = RewriteFormulaForPaste("=" + newRaw, dr, dc); } catch { }
                            }
                        }
                        string? oldRaw = _sheet.GetRaw(r, c);
                        if (oldRaw == newRaw) continue;
                        _sheet.SetRaw(r, c, newRaw);
                        edits.Add((r, c, oldRaw, newRaw));
                        foreach (var ac in _sheet.RecalculateDirty(r, c)) affected.Add(ac);
                    }
                }
            }
            finally { _suppressRecord = false; }
            if (edits.Count > 0)
            {
                _undo.RecordBulk(edits);
                try { RefreshDirtyOrFull(affected); } catch { }
                try { UpdateAiMenuItemsState(); } catch { }
            }
        }

        // --- Find / Replace ---
        private void OpenFindDialog(bool focusReplace)
        {
            if (_findDlg == null || _findDlg.IsDisposed)
            {
                _findDlg = new FindReplaceForm();
                _findDlg.FindNextClicked += (s, e) => DoFindNextFromDialog();
                _findDlg.ReplaceClicked += (s, e) => DoReplaceFromDialog();
                _findDlg.ReplaceAllClicked += (s, e) => DoReplaceAllFromDialog();
            }
            _findDlg.FindText = _lastFind;
            _findDlg.ReplaceText = _lastReplace;
            _findDlg.MatchCase = _lastMatchCase;

            _findDlg.Show(this);
            _findDlg.BringToFront();
        }

        private StringComparison FindComparison(bool matchCase) => matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

        private bool ContainsFind(string? raw, string needle, bool matchCase)
        {
            raw ??= string.Empty;
            if (string.IsNullOrEmpty(needle)) return false;
            return raw.IndexOf(needle, FindComparison(matchCase)) >= 0;
        }

        private bool DoFindNext(string needle, bool matchCase)
        {
            if (string.IsNullOrEmpty(needle)) return false;
            int startR = _findPosRow;
            int startC = _findPosCol;
            // Start after current cell
            int r = startR;
            int c = startC + 1;
            for (int pass = 0; pass < 2; pass++)
            {
                for (; r < _sheet.Rows; r++, c = 0)
                {
                    for (; c < _sheet.Columns; c++)
                    {
                        var raw = _sheet.GetRaw(r, c);
                        if (ContainsFind(raw, needle, matchCase))
                        {
                            _findPosRow = r; _findPosCol = c;
                            grid.CurrentCell = grid[c, r];
                            grid.FirstDisplayedScrollingRowIndex = Math.Max(0, r - 3);
                            UpdateStatus();
                            return true;
                        }
                    }
                }
                // wrap
                r = 0; c = 0;
            }
            return false;
        }

        private void DoFindNextFromDialog()
        {
            if (_findDlg == null) return;
            _lastFind = _findDlg.FindText;
            _lastMatchCase = _findDlg.MatchCase;
            if (grid.CurrentCell != null)
            {
                _findPosRow = grid.CurrentCell.RowIndex;
                _findPosCol = grid.CurrentCell.ColumnIndex;
            }
            if (!DoFindNext(_lastFind, _lastMatchCase))
            {
                System.Media.SystemSounds.Beep.Play();
            }
        }

        private string ReplaceAllInString(string source, string find, string repl, bool matchCase)
        {
            if (string.IsNullOrEmpty(find)) return source;
            var comp = FindComparison(matchCase);
            int index = source.IndexOf(find, comp);
            if (index < 0) return source;
            var sb = new System.Text.StringBuilder();
            int last = 0;
            while (index >= 0)
            {
                sb.Append(source, last, index - last);
                sb.Append(repl);
                last = index + find.Length;
                index = source.IndexOf(find, last, comp);
            }
            sb.Append(source, last, source.Length - last);
            return sb.ToString();
        }

        private void DoReplaceFromDialog()
        {
            if (_findDlg == null) return;
            _lastFind = _findDlg.FindText;
            _lastReplace = _findDlg.ReplaceText;
            _lastMatchCase = _findDlg.MatchCase;
            if (string.IsNullOrEmpty(_lastFind)) return;

            // If current cell doesn't have a match, find next
            if (grid.CurrentCell == null || !ContainsFind(_sheet.GetRaw(grid.CurrentCell.RowIndex, grid.CurrentCell.ColumnIndex), _lastFind, _lastMatchCase))
            {
                if (!DoFindNext(_lastFind, _lastMatchCase)) { System.Media.SystemSounds.Beep.Play(); return; }
            }

            int r = grid.CurrentCell!.RowIndex;
            int c = grid.CurrentCell!.ColumnIndex;
            var raw = _sheet.GetRaw(r, c) ?? string.Empty;
            var comp = FindComparison(_lastMatchCase);
            int idx = raw.IndexOf(_lastFind, comp);
            if (idx >= 0)
            {
                string newRaw = raw.Substring(0, idx) + _lastReplace + raw.Substring(idx + _lastFind.Length);
                var oldRaw = raw;
                _suppressRecord = true; try { _sheet.SetRaw(r, c, newRaw); } finally { _suppressRecord = false; }
                _undo.RecordSet(r, c, oldRaw, newRaw);
                var affected = _sheet.RecalculateDirty(r, c);
                RefreshDirtyOrFull(affected);

                // Move to next
                _findPosRow = r; _findPosCol = c;
                DoFindNext(_lastFind, _lastMatchCase);
                try { UpdateAiMenuItemsState(); } catch { }
            }
        }

        private void DoReplaceAllFromDialog()
        {
            if (_findDlg == null) return;
            _lastFind = _findDlg.FindText;
            _lastReplace = _findDlg.ReplaceText;
            _lastMatchCase = _findDlg.MatchCase;
            if (string.IsNullOrEmpty(_lastFind)) return;

            bool any = false;
            var affectedAll = new HashSet<(int r, int c)>();
            for (int r = 0; r < _sheet.Rows; r++)
            {
                for (int c = 0; c < _sheet.Columns; c++)
                {
                    var raw = _sheet.GetRaw(r, c) ?? string.Empty;
                    if (!ContainsFind(raw, _lastFind, _lastMatchCase)) continue;
                    string replaced = ReplaceAllInString(raw, _lastFind, _lastReplace, _lastMatchCase);
                    if (replaced != raw)
                    {
                        any = true;
                        var oldRaw = raw;
                        _suppressRecord = true; try { _sheet.SetRaw(r, c, replaced); } finally { _suppressRecord = false; }
                        _undo.RecordSet(r, c, oldRaw, replaced);
                        foreach (var ac in _sheet.RecalculateDirty(r, c)) affectedAll.Add(ac);
                    }
                }
            }
            if (any)
            {
                RefreshDirtyOrFull(affectedAll);
                try { UpdateAiMenuItemsState(); } catch { }
            }
        }

        private void RecalculateAll()
        {
            try
            {
                RefreshGridValues();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Error during calculation: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ToggleInlineSuggestions(bool enabled)
        {
            _aiInlineEnabled = enabled;
            if (!enabled)
            {
                HideGhostSuggestions();
                CancelInlineRequest();
            }
            else
            {
                ScheduleInlineSuggestion();
            }
            SyncAiMenuState();
        }

        private void CreateGhostUI()
        {
            _aiGhostPanel = new Panel
            {
                Visible = false,
                BackColor = System.Drawing.Color.FromArgb(16, 0, 128, 255),
                BorderStyle = BorderStyle.FixedSingle,
                Height = 120,
                Width = 240
            };
            _aiGhostStatus = new Label { Dock = DockStyle.Top, Height = 18, ForeColor = System.Drawing.Color.Gray, Text = string.Empty };
            _aiGhostApplyBtn = new Button { Text = "Apply", Dock = DockStyle.Bottom, Height = 22 };
            _aiGhostDismissBtn = new Button { Text = "Dismiss", Dock = DockStyle.Bottom, Height = 22 };
            _aiGhostList = new ListBox
            {
                Dock = DockStyle.Fill,
                ForeColor = System.Drawing.Color.Gray
            };
            _aiGhostPanel.Controls.Add(_aiGhostList);
            _aiGhostPanel.Controls.Add(_aiGhostApplyBtn);
            _aiGhostPanel.Controls.Add(_aiGhostDismissBtn);
            _aiGhostPanel.Controls.Add(_aiGhostStatus);
            grid.Controls.Add(_aiGhostPanel);
            grid.Scroll += (s, e) => HideGhostSuggestions();
            if (_aiGhostList != null)
            {
                _aiGhostList.DoubleClick += (s, e) => AcceptInlineSuggestion();
                _aiGhostList.KeyDown += (s, e) =>
                {
                    if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
                    {
                        e.Handled = true; AcceptInlineSuggestion();
                    }
                };
            }
            if (_aiGhostApplyBtn != null) _aiGhostApplyBtn.Click += (s, e) => AcceptInlineSuggestion();
            if (_aiGhostDismissBtn != null) _aiGhostDismissBtn.Click += (s, e) => HideGhostSuggestions();
        }

        private void ShowGhostSuggestions(int startRow, int col, string[] items)
        {
            if (_aiGhostPanel == null || _aiGhostList == null) return;
            if (_aiGhostStatus != null) _aiGhostStatus.Text = string.Empty;
            _aiGhostStartRow = startRow;
            _aiGhostCol = col;
            _aiGhostList.Items.Clear();
            foreach (var it in items) _aiGhostList.Items.Add(it);
            if (_aiGhostApplyBtn != null) _aiGhostApplyBtn.Enabled = items.Length > 0;
            var rect = grid.GetCellDisplayRectangle(col, startRow, true);
            rect.Y = rect.Bottom; // below the start cell
            rect.Height = Math.Min(120, items.Length * 18 + 8);
            rect.Width = Math.Max(rect.Width, 180);
            _aiGhostPanel.Bounds = rect;
            _aiGhostPanel.Visible = true;
            _aiGhostPanel.BringToFront();
            try { _aiGhostList.SelectedIndex = _aiGhostList.Items.Count > 0 ? 0 : -1; } catch { }
            UpdateAiMenuItemsState();
        }

        private void ShowGhostStatus(int startRow, int col, string message)
        {
            if (_aiGhostPanel == null) return;
            if (_aiGhostStatus != null) _aiGhostStatus.Text = message;
            if (_aiGhostList != null) _aiGhostList.Items.Clear();
            if (_aiGhostApplyBtn != null) _aiGhostApplyBtn.Enabled = false;
            var rect = grid.GetCellDisplayRectangle(col, startRow, true);
            rect.Y = rect.Bottom; // below start cell
            rect.Height = 40;
            rect.Width = Math.Max(rect.Width, 180);
            _aiGhostPanel.Bounds = rect;
            _aiGhostPanel.Visible = true;
            _aiGhostPanel.BringToFront();
            UpdateAiMenuItemsState();
        }

        private void HideGhostSuggestions()
        {
            if (_aiGhostPanel != null) _aiGhostPanel.Visible = false;
            _aiGhostStartRow = -1; _aiGhostCol = -1;
            UpdateAiMenuItemsState();
        }

        private string GetRecentColumnKey(int startRow, int col)
        {
            var sb = new System.Text.StringBuilder();
            int count = 0;
            for (int r = startRow - 1; r >= 0 && count < 3; r--)
            {
                var raw = _sheet.GetRaw(r, col);
                if (string.IsNullOrWhiteSpace(raw)) break;
                if (sb.Length > 0) sb.Append("||");
                sb.Append(raw.Trim().ToLowerInvariant());
                count++;
            }
            return sb.ToString();
        }

        private void UpdateAiMenuItemsState()
        {
            try
            {
                // Enable Accept when a ghost suggestion is visible
                aiAcceptInlineToolStripMenuItem.Enabled = _aiGhostPanel != null && _aiGhostPanel.Visible;
                // Global enable/disable
                bool providerReady = ProviderReady();
                aiGenerateFillToolStripMenuItem.Enabled = _settings.AiEnabled && providerReady;
                aiToggleChatPaneToolStripMenuItem.Enabled = _settings.AiEnabled;
                aiEnableInlineToolStripMenuItem.Enabled = _settings.AiEnabled && providerReady;
                if (!_settings.AiEnabled || !providerReady)
                {
                    aiEnableInlineToolStripMenuItem.Checked = false;
                }
                // Undo/Redo availability
                try
                {
                    undoToolStripMenuItem.Enabled = _undo.CanUndo;
                    redoToolStripMenuItem.Enabled = _undo.CanRedo;
                }
                catch { }
            }
            catch { }
        }

        // --- Selection bounds helper ---
        private bool TryGetSelectionBounds(out int top, out int left, out int bottom, out int right)
        {
            if (grid.CurrentCell == null)
            {
                top = left = bottom = right = 0; return false;
            }
            top = grid.CurrentCell.RowIndex; left = grid.CurrentCell.ColumnIndex; bottom = top; right = left;
            var cells = grid.SelectedCells;
            if (cells != null && cells.Count > 1)
            {
                foreach (DataGridViewCell cell in cells)
                {
                    if (cell == null) continue;
                    top = Math.Min(top, cell.RowIndex);
                    left = Math.Min(left, cell.ColumnIndex);
                    bottom = Math.Max(bottom, cell.RowIndex);
                    right = Math.Max(right, cell.ColumnIndex);
                }
            }
            return true;
        }

        // --- Clipboard outline drawing ---
        private void ShowClipboardOutline(int top, int left, int bottom, int right, bool isCut)
        {
            _outlineTop = Math.Max(0, top);
            _outlineLeft = Math.Max(0, left);
            _outlineBottom = Math.Max(_outlineTop, bottom);
            _outlineRight = Math.Max(_outlineLeft, right);
            _clipboardOutlineIsCut = isCut;
            _clipboardOutlineVisible = true;
            try { grid.Invalidate(); } catch { }
        }

        private void ClearClipboardOutline()
        {
            if (!_clipboardOutlineVisible) return;
            _clipboardOutlineVisible = false;
            try { grid.Invalidate(); } catch { }
        }

        private void Grid_Paint(PaintEventArgs e)
        {
            // Clipboard outline
            if (_clipboardOutlineVisible)
            {
                try
                {
                    var tl = grid.GetCellDisplayRectangle(_outlineLeft, _outlineTop, true);
                    var br = grid.GetCellDisplayRectangle(_outlineRight, _outlineBottom, true);
                    if (tl.Width > 0 && tl.Height > 0 && br.Width > 0 && br.Height > 0)
                    {
                        var rect = System.Drawing.Rectangle.FromLTRB(tl.Left, tl.Top, br.Right - 1, br.Bottom - 1);
                        using var pen = new System.Drawing.Pen(System.Drawing.Color.FromArgb(0, 120, 215), 2f);
                        if (_clipboardOutlineIsCut) pen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;
                        e.Graphics.DrawRectangle(pen, rect);
                    }
                }
                catch { }
            }
            // Drag-fill handle & preview
            try
            {
                if (TryGetSelectionBounds(out int top, out int left, out int bottom, out int right))
                {
                    var brCell = grid.GetCellDisplayRectangle(right, bottom, true);
                    if (brCell.Width > 0 && brCell.Height > 0)
                    {
                        var handle = new System.Drawing.Rectangle(brCell.Right - 7, brCell.Bottom - 7, 7, 7);
                        _dragHandleRect = handle;
                        using var brush = new System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(0, 120, 215));
                        e.Graphics.FillRectangle(brush, handle);
                    }
                }
                if (_dragFillPreview)
                {
                    var tl = grid.GetCellDisplayRectangle(_previewLeft, _previewTop, true);
                    var br = grid.GetCellDisplayRectangle(_previewRight, _previewBottom, true);
                    if (tl.Width > 0 && tl.Height > 0 && br.Width > 0 && br.Height > 0)
                    {
                        var rect = System.Drawing.Rectangle.FromLTRB(tl.Left, tl.Top, br.Right - 1, br.Bottom - 1);
                        using var pen = new System.Drawing.Pen(System.Drawing.Color.FromArgb(0, 120, 215));
                        pen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dot;
                        e.Graphics.DrawRectangle(pen, rect);
                    }
                }
            }
            catch { }
        }

        private void Grid_MouseDown(object? sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;
            if (_dragHandleRect.Contains(e.Location))
            {
                if (!TryGetSelectionBounds(out _selTop, out _selLeft, out _selBottom, out _selRight)) return;
                _dragFillActive = true;
                _dragFillPreview = false;
                try { grid.Capture = true; } catch { }
            }
        }

        private void Grid_MouseMove(object? sender, MouseEventArgs e)
        {
            try { if (_dragHandleRect.Contains(e.Location)) grid.Cursor = Cursors.Cross; else if (!_dragFillActive) grid.Cursor = Cursors.Default; } catch { }
            if (!_dragFillActive) return;
            var hit = grid.HitTest(e.X, e.Y);
            if (hit.RowIndex < 0 || hit.ColumnIndex < 0) { _dragFillPreview = false; try { grid.Invalidate(); } catch { } return; }
            int r = Math.Min(Math.Max(0, hit.RowIndex), _sheet.Rows - 1);
            int c = Math.Min(Math.Max(0, hit.ColumnIndex), _sheet.Columns - 1);
            int pt = Math.Min(_selTop, r);
            int pl = Math.Min(_selLeft, c);
            int pb = Math.Max(_selBottom, r);
            int pr = Math.Max(_selRight, c);
            bool extends = (pb > _selBottom) || (pr > _selRight) || (pt < _selTop) || (pl < _selLeft);
            _dragFillPreview = extends;
            if (extends)
            {
                _previewTop = pt; _previewLeft = pl; _previewBottom = pb; _previewRight = pr;
            }
            try { grid.Invalidate(); } catch { }
        }

        private void Grid_MouseUp(object? sender, MouseEventArgs e)
        {
            if (!_dragFillActive) return;
            try { grid.Capture = false; } catch { }
            _dragFillActive = false;
            if (!_dragFillPreview)
            {
                _dragFillPreview = false; try { grid.Invalidate(); } catch { } return;
            }
            int t = _previewTop, l = _previewLeft, b = _previewBottom, r = _previewRight;
            ApplyDragFill(t, l, b, r);
            try { grid.Invalidate(); } catch { }
        }

        private void Grid_MouseDoubleClick(object? sender, MouseEventArgs e)
        {
            if (!_dragHandleRect.Contains(e.Location)) return;
            // Determine reference column for boundary detection (prefer the input column to the left)
            if (!TryGetSelectionBounds(out _selTop, out _selLeft, out _selBottom, out _selRight)) return;
            int refCol = _selLeft > 0 ? _selLeft - 1 : _selLeft;
            int r = _selBottom + 1;
            while (r < _sheet.Rows)
            {
                var raw = _sheet.GetRaw(r, refCol) ?? string.Empty;
                if (string.IsNullOrWhiteSpace(raw)) break;
                r++;
            }
            int targetBottom = r - 1;
            if (targetBottom <= _selBottom) return; // nothing to extend
            ApplyDragFill(_selTop, _selLeft, targetBottom, _selRight);
            try { grid.Invalidate(); } catch { }
        }

        private void ApplyDragFill(int t, int l, int b, int r)
        {
            var edits = new List<(int row, int col, string? oldRaw, string? newRaw)>();
            var affected = new HashSet<(int r, int c)>();
            _suppressRecord = true;
            try
            {
                // Determine extension orientation for series detection
                bool downOnly = (l == _selLeft && r == _selRight && t == _selTop && b > _selBottom);
                bool rightOnly = (t == _selTop && b == _selBottom && l == _selLeft && r > _selRight);
                if (downOnly)
                {
                    // Column-wise series detection on numeric values; fallback to pattern copy
                    int baseTop = _selTop, baseBottom = _selBottom;
                    for (int c = _selLeft; c <= _selRight; c++)
                    {
                        bool hasFormula = false;
                        var nums = new List<double>();
                        for (int rr = baseTop; rr <= baseBottom; rr++)
                        {
                            var raw = _sheet.GetRaw(rr, c) ?? string.Empty;
                            if (!string.IsNullOrEmpty(raw) && raw.StartsWith("=", StringComparison.Ordinal)) { hasFormula = true; break; }
                            var v = _sheet.GetValue(rr, c);
                            if (v.Error == null && v.Number is double d) nums.Add(d);
                        }
                        double? step = null; double? last = null;
                        if (!hasFormula && nums.Count >= 2)
                        {
                            // Use last two numeric values to compute step
                            last = nums[^1]; step = nums[^1] - nums[^2];
                        }
                        for (int rr = baseBottom + 1; rr <= b; rr++)
                        {
                            string? oldRaw = _sheet.GetRaw(rr, c);
                            string newRaw;
                            if (step.HasValue && last.HasValue)
                            {
                                double k = rr - baseBottom;
                                double val = last.Value + step.Value * k;
                                newRaw = val.ToString(System.Globalization.CultureInfo.InvariantCulture);
                            }
                            else
                            {
                                // Fallback: copy last row with rewrite
                                int sr = baseBottom; int sc = c; int dr = rr - sr; int dc = 0;
                                var src = _sheet.GetRaw(sr, sc) ?? string.Empty;
                                newRaw = src;
                                if (!string.IsNullOrEmpty(newRaw))
                                {
                                    if (newRaw.StartsWith("=", StringComparison.Ordinal))
                                        try { newRaw = RewriteFormulaForPaste(newRaw, dr, dc); } catch { }
                                    else if (LooksLikeCellRefOrRange(newRaw))
                                        try { newRaw = RewriteFormulaForPaste("=" + newRaw, dr, dc); } catch { }
                                }
                            }
                            if (oldRaw == newRaw) continue;
                            _sheet.SetRaw(rr, c, newRaw);
                            edits.Add((rr, c, oldRaw, newRaw));
                            foreach (var ac in _sheet.RecalculateDirty(rr, c)) affected.Add(ac);
                        }
                    }
                }
                else if (rightOnly)
                {
                    // Row-wise series detection; fallback to pattern copy
                    int baseLeft = _selLeft, baseRight = _selRight;
                    for (int rrow = _selTop; rrow <= _selBottom; rrow++)
                    {
                        bool hasFormula = false;
                        var nums = new List<double>();
                        for (int cc = baseLeft; cc <= baseRight; cc++)
                        {
                            var raw = _sheet.GetRaw(rrow, cc) ?? string.Empty;
                            if (!string.IsNullOrEmpty(raw) && raw.StartsWith("=", StringComparison.Ordinal)) { hasFormula = true; break; }
                            var v = _sheet.GetValue(rrow, cc);
                            if (v.Error == null && v.Number is double d) nums.Add(d);
                        }
                        double? step = null; double? last = null;
                        if (!hasFormula && nums.Count >= 2)
                        {
                            last = nums[^1]; step = nums[^1] - nums[^2];
                        }
                        for (int cc = baseRight + 1; cc <= r; cc++)
                        {
                            string? oldRaw = _sheet.GetRaw(rrow, cc);
                            string newRaw;
                            if (step.HasValue && last.HasValue)
                            {
                                double k = cc - baseRight;
                                double val = last.Value + step.Value * k;
                                newRaw = val.ToString(System.Globalization.CultureInfo.InvariantCulture);
                            }
                            else
                            {
                                int sr = rrow; int sc = baseRight; int dr = 0; int dc = cc - sc;
                                var src = _sheet.GetRaw(sr, sc) ?? string.Empty;
                                newRaw = src;
                                if (!string.IsNullOrEmpty(newRaw))
                                {
                                    if (newRaw.StartsWith("=", StringComparison.Ordinal))
                                        try { newRaw = RewriteFormulaForPaste(newRaw, dr, dc); } catch { }
                                    else if (LooksLikeCellRefOrRange(newRaw))
                                        try { newRaw = RewriteFormulaForPaste("=" + newRaw, dr, dc); } catch { }
                                }
                            }
                            if (oldRaw == newRaw) continue;
                            _sheet.SetRaw(rrow, cc, newRaw);
                            edits.Add((rrow, cc, oldRaw, newRaw));
                            foreach (var ac in _sheet.RecalculateDirty(rrow, cc)) affected.Add(ac);
                        }
                    }
                }
                else
                {
                    // Fallback: repeating pattern with rewrite for formulas/ranges
                    int baseH = _selBottom - _selTop + 1;
                    int baseW = _selRight - _selLeft + 1;
                    for (int rr = t; rr <= b; rr++)
                    {
                        for (int cc = l; cc <= r; cc++)
                        {
                            if (rr >= _selTop && rr <= _selBottom && cc >= _selLeft && cc <= _selRight) continue;
                            int sr = _selTop + ((rr - _selTop) % baseH + baseH) % baseH;
                            int sc = _selLeft + ((cc - _selLeft) % baseW + baseW) % baseW;
                            string src = _sheet.GetRaw(sr, sc) ?? string.Empty;
                            string newRaw = src;
                            int dr = rr - sr; int dc = cc - sc;
                            if (!string.IsNullOrEmpty(newRaw))
                            {
                                if (newRaw.StartsWith("=", StringComparison.Ordinal))
                                {
                                    try { newRaw = RewriteFormulaForPaste(newRaw, dr, dc); } catch { }
                                }
                                else if (LooksLikeCellRefOrRange(newRaw))
                                {
                                    try { newRaw = RewriteFormulaForPaste("=" + newRaw, dr, dc); } catch { }
                                }
                            }
                            string? oldRaw = _sheet.GetRaw(rr, cc);
                            if (oldRaw == newRaw) continue;
                            _sheet.SetRaw(rr, cc, newRaw);
                            edits.Add((rr, cc, oldRaw, newRaw));
                            foreach (var ac in _sheet.RecalculateDirty(rr, cc)) affected.Add(ac);
                        }
                    }
                }
            }
            finally { _suppressRecord = false; }
            _dragFillPreview = false;
            if (edits.Count > 0)
            {
                _undo.RecordBulk(edits);
                try { RefreshDirtyOrFull(affected); } catch { }
                try { UpdateAiMenuItemsState(); } catch { }
            }
        }

        private void ScheduleInlineSuggestion()
        {
            if (!_aiInlineEnabled) return;
            if (!_settings.AiEnabled || !ProviderReady()) return;
            _aiInlineTimer.Stop();
            _aiInlineTimer.Start();
        }

        private void CancelInlineRequest()
        {
            try { _aiInlineCts?.Cancel(); } catch { }
            _aiInlineCts = null;
        }

        private async System.Threading.Tasks.Task TriggerInlineSuggestionAsync()
        {
            if (!_aiInlineEnabled) return;
            if (!_settings.AiEnabled || !ProviderReady()) return;
            if (grid.CurrentCell == null) return;
            if (grid.IsCurrentCellInEditMode) return; // keep it simple: suggest when not editing
            int r = grid.CurrentCell.RowIndex;
            int c = grid.CurrentCell.ColumnIndex;
            // Gate: current cell empty and at least 2 non-empty cells above in same column
            var cur = _sheet.GetRaw(r, c) ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(cur)) { HideGhostSuggestions(); return; }
            int nonEmpty = 0;
            int scan = r - 1;
            while (scan >= 0 && nonEmpty < 3)
            {
                var raw = _sheet.GetRaw(scan, c);
                if (string.IsNullOrWhiteSpace(raw)) break;
                nonEmpty++; scan--;
            }
            if (nonEmpty < 2) { HideGhostSuggestions(); return; }

            // Determine title (use row 0 if present, else the first empty above the list)
            string? title = _sheet.GetRaw(0, c);
            if (string.IsNullOrWhiteSpace(title) && scan >= 0)
            {
                var maybe = _sheet.GetRaw(scan, c);
                if (!string.IsNullOrWhiteSpace(maybe)) title = maybe;
            }

            int availableRows = _sheet.Rows - r;
            int n = Math.Max(1, Math.Min(availableRows, 5));
            string sheetName = _sheetNames.Count > _activeSheetIndex ? _sheetNames[_activeSheetIndex] : "Sheet";
            // Build small cache key from recent items and title
            string recentKey = GetRecentColumnKey(r, c);
            string cacheKey = $"{sheetName}|c{c}|{recentKey}|{(title ?? string.Empty).Trim().ToLowerInvariant()}|n{n}";
            if (_aiInlineCache.TryGetValue(cacheKey, out var cached) && cached.Length > 0)
            {
                ShowGhostSuggestions(r, c, cached);
                return;
            }
            var ctx = new AIContext
            {
                SheetName = sheetName,
                StartRow = r,
                StartCol = c,
                Rows = n,
                Cols = 1,
                Title = title,
                Prompt = "continue the list"
            };

            CancelInlineRequest();
            _aiInlineCts = new System.Threading.CancellationTokenSource();
            var ct = _aiInlineCts.Token;
            try
            {
                using var timeout = new System.Threading.CancellationTokenSource(TimeSpan.FromSeconds(5));
                using var linked = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(ct, timeout.Token);
                ShowGhostStatus(r, c, "Generating…");
                var res = await _aiProvider.GenerateFillAsync(ctx, linked.Token).ConfigureAwait(true);
                // Flatten to 1D and trim any leading duplicates of the items above
                var items = new System.Collections.Generic.List<string>();
                foreach (var row in res.Cells)
                {
                    if (row.Length > 0) items.Add(row[0] ?? string.Empty);
                }
                if (items.Count == 0) { HideGhostSuggestions(); return; }
                var filtered = TrimLeadingSeenThenDedupe(r, c, items);
                if (filtered.Count == 0)
                {
                    ShowGhostStatus(r, c, "No suggestions");
                    return;
                }
                var arr = filtered.ToArray();
                _aiInlineCache[cacheKey] = arr;
                ShowGhostSuggestions(r, c, arr);
            }
            catch (OperationCanceledException)
            {
                // ignore
            }
            catch
            {
                HideGhostSuggestions();
            }
        }

        private static string NormalizeItem(string s)
        {
            return (s ?? string.Empty).Trim().ToLowerInvariant();
        }

        private System.Collections.Generic.List<string> CollectColumnItemsAboveList(int startRow, int col)
        {
            var list = new System.Collections.Generic.List<string>();
            int r = startRow - 1;
            while (r >= 0)
            {
                var raw = _sheet.GetRaw(r, col) ?? string.Empty;
                if (string.IsNullOrWhiteSpace(raw)) break; // stop at first gap
                list.Add(NormalizeItem(raw));
                r--;
            }
            list.Reverse(); // top-to-bottom order
            return list;
        }

        private System.Collections.Generic.List<string> TrimLeadingSeenThenDedupe(int startRow, int col, System.Collections.Generic.IEnumerable<string> items)
        {
            var aboveList = CollectColumnItemsAboveList(startRow, col);
            var aboveSet = new System.Collections.Generic.HashSet<string>(aboveList, System.StringComparer.OrdinalIgnoreCase);
            var array = new System.Collections.Generic.List<string>(items);
            int i = 0;
            while (i < array.Count && aboveSet.Contains(NormalizeItem(array[i]))) i++;
            var result = new System.Collections.Generic.List<string>();
            var dedupe = new System.Collections.Generic.HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
            for (; i < array.Count; i++)
            {
                var it = array[i];
                var norm = NormalizeItem(it);
                if (string.IsNullOrEmpty(norm)) continue;
                if (dedupe.Add(norm)) result.Add(it);
            }
            return result;
        }

        private void AcceptInlineSuggestion()
        {
            if (_aiGhostPanel == null || !_aiGhostPanel.Visible || _aiGhostStartRow < 0 || _aiGhostCol < 0) return;
            if (_aiGhostList == null) return;
            int count = _aiGhostList.Items.Count;
            if (count <= 0) return;
            var cells = new string[count][];
            for (int i = 0; i < count; i++) cells[i] = new[] { _aiGhostList.Items[i]?.ToString() ?? string.Empty };
            int r = _aiGhostStartRow;
            int c = _aiGhostCol;
            HideGhostSuggestions();
            ApplyCellsWithUndo(r, c, cells);
        }

        // --- AI: Chat Assistant ---
        private void OpenChatAssistant()
        {
            if (!_settings.AiEnabled) { MessageBox.Show(this, "AI is disabled in Settings.", "AI", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            if (_chatDockHost == null || _chatDockSplitter == null)
            {
                try { CreateDockedChatPane(); } catch { }
            }
            if (_chatDockHost != null)
            {
                _chatDockHost.Visible = true;
                if (_chatDockSplitter != null) _chatDockSplitter.Visible = true;
                try { _chatPane?.FocusInput(); } catch { }
            }
        }

        // --- AI: Schema Fill (single-shot values-only) ---
        private async System.Threading.Tasks.Task FillSelectedFromSchemaAsync()
        {
            if (!_settings.AiEnabled) { try { MessageBox.Show(this, "AI is disabled in Settings.", "AI", MessageBoxButtons.OK, MessageBoxIcon.Information); } catch { } return; }
            if (!ProviderReady()) { try { MessageBox.Show(this, "No AI provider is configured.", "AI", MessageBoxButtons.OK, MessageBoxIcon.Information); } catch { } return; }
            var ctx = BuildPlannerContext();
            // Gate to values-only
            ctx.AllowedCommands = new[] { "set_values" };
            string prompt = "Fill the selected range from the input column using the headers as schema. Use set_values only. Do not modify headers or the input column. Return a single set_values sized exactly to the selection.";
            int timeoutSec = 30;
            try { var s = Environment.GetEnvironmentVariable("AI_PLAN_TIMEOUT_SEC"); if (!string.IsNullOrWhiteSpace(s)) timeoutSec = Math.Max(5, int.Parse(s)); } catch { }
            SpreadsheetApp.Core.AI.AIPlan plan;
            using (var cts = new System.Threading.CancellationTokenSource(System.TimeSpan.FromSeconds(timeoutSec)))
            {
                try { plan = await _chatPlanner.PlanAsync(ctx, prompt, cts.Token).ConfigureAwait(true); }
                catch (Exception ex)
                {
                    try { MessageBox.Show(this, $"Planning failed: {ex.Message}", "AI", MessageBoxButtons.OK, MessageBoxIcon.Error); } catch { }
                    return;
                }
            }
            // Sanitize to current selection bounds before apply
            int minR, minC, maxR, maxC;
            TryGetSelectionBounds(out minR, out minC, out maxR, out maxC);
            try { plan = SanitizePlanToBounds(plan, minR, minC, maxR, maxC); } catch { }
            if (plan.Commands.Count == 0) { try { MessageBox.Show(this, "No changes suggested.", "AI", MessageBoxButtons.OK, MessageBoxIcon.Information); } catch { } return; }
            ApplyPlan(plan);
        }

        private async System.Threading.Tasks.Task BatchFillSelectedFromSchemaAsync()
        {
            if (!_settings.AiEnabled) { try { MessageBox.Show(this, "AI is disabled.", "AI", MessageBoxButtons.OK, MessageBoxIcon.Information); } catch { } return; }
            if (!ProviderReady()) { try { MessageBox.Show(this, "No AI provider configured.", "AI", MessageBoxButtons.OK, MessageBoxIcon.Information); } catch { } return; }

            var ctx = BuildPlannerContext();
            ctx.AllowedCommands = new[] { "set_values" };

            // Detect input column values
            int inputCol = Math.Max(0, ctx.StartCol - 1);
            var inputs = new List<string>();
            for (int r = ctx.StartRow; r < ctx.StartRow + ctx.Rows; r++)
            {
                var raw = _sheet.GetRaw(r, inputCol) ?? string.Empty;
                inputs.Add(raw);
            }

            int totalRows = inputs.Count;
            if (totalRows <= 40)
            {
                // Small enough for single-shot
                await FillSelectedFromSchemaAsync();
                return;
            }

            // Confirm batch fill
            var filler = new BatchSchemaFiller(_chatPlanner);
            int batchCount = filler.EstimateBatchCount(totalRows);
            var confirm = MessageBox.Show(this,
                $"This will fill {totalRows} rows in {batchCount} batches.\nProceed?",
                "Batch Schema Fill", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (confirm != DialogResult.OK) return;

            string prompt = "Fill the selected range from the input column using the headers as schema. Use set_values only. Do not modify headers or the input column. Return a single set_values sized exactly to the batch selection.";

            using var cts = new System.Threading.CancellationTokenSource();
            // Simple progress via title bar
            string origTitle = Text;
            var results = await filler.FillAsync(ctx, prompt, inputs.ToArray(), ctx.StartRow,
                progress =>
                {
                    try { BeginInvoke(new Action(() => Text = $"{origTitle} — Batch {progress.CompletedBatches}/{progress.TotalBatches}")); } catch { }
                },
                cts.Token);

            Text = origTitle;

            // Apply successful batches
            int applied = 0;
            foreach (var result in results)
            {
                if (result.Plan != null)
                {
                    try
                    {
                        int minR, minC, maxR, maxC;
                        minR = result.StartRow; minC = ctx.StartCol;
                        maxR = result.StartRow + result.RowCount - 1; maxC = ctx.StartCol + ctx.Cols - 1;
                        var sanitized = SanitizePlanToBounds(result.Plan, minR, minC, maxR, maxC);
                        ApplyPlan(sanitized);
                        applied++;
                    }
                    catch { }
                }
            }

            int failed = results.Count(r => !r.Success);
            if (failed > 0)
                MessageBox.Show(this, $"Completed: {applied} batches applied, {failed} failed.", "Batch Fill", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private AIContext BuildPlannerContext()
        {
            int sr = grid.CurrentCell?.RowIndex ?? 0;
            int sc = grid.CurrentCell?.ColumnIndex ?? 0;
            int rowsHint = 5, colsHint = 1;
            bool hadSelection = false;
            try
            {
                // Selection values (bounded)
                if (grid.SelectedCells != null && grid.SelectedCells.Count > 1)
                {
                    int minR = int.MaxValue, maxR = -1, minC = int.MaxValue, maxC = -1;
                    foreach (DataGridViewCell cell in grid.SelectedCells)
                    {
                        if (cell == null) continue;
                        if (cell.RowIndex < minR) minR = cell.RowIndex;
                        if (cell.RowIndex > maxR) maxR = cell.RowIndex;
                        if (cell.ColumnIndex < minC) minC = cell.ColumnIndex;
                        if (cell.ColumnIndex > maxC) maxC = cell.ColumnIndex;
                    }
                    if (minR <= maxR && minC <= maxC)
                    {
                        hadSelection = true;
                        // Update context start to selection top-left
                        sr = minR; sc = minC;
                        int rows = Math.Min(40, maxR - minR + 1);
                        int cols = Math.Min(10, maxC - minC + 1);
                        rowsHint = rows; colsHint = cols;
                        var sel = new string[rows][];
                        for (int r = 0; r < rows; r++)
                        {
                            sel[r] = new string[cols];
                            for (int c = 0; c < cols; c++)
                            {
                                int rr = minR + r; int cc = minC + c;
                                sel[r][c] = _sheet.GetDisplay(rr, cc);
                            }
                        }
                        // We'll set this after creating ctx
                        // temp hold
                        _lastSelVals = sel;
                    }
                }
            }
            catch { }
            // Recompute title after possibly moving sr/sc
            string? title = sr > 0 ? _sheet.GetRaw(sr - 1, sc) : null;
            var ctx = new AIContext
            {
                SheetName = _sheetNames.Count > _activeSheetIndex ? _sheetNames[_activeSheetIndex] : "Sheet",
                StartRow = sr,
                StartCol = sc,
                Rows = rowsHint,
                Cols = colsHint,
                Title = title,
                Prompt = string.Empty
            };
            try
            {
                if (hadSelection && _lastSelVals != null)
                {
                    ctx.SelectionValues = _lastSelVals;
                }
            }
            catch { }
            try
            {
                // Nearby values window
                int winRows = 20, winCols = 10;
                int startR = Math.Max(0, sr - 1);  // include header row above selection
                int startC = Math.Max(0, sc - 1);  // include one column to the left for context
                int rows = Math.Min(winRows, _sheet.Rows - startR);
                int cols = Math.Min(winCols, _sheet.Columns - startC);
                var near = new string[rows][];
                for (int r = 0; r < rows; r++)
                {
                    near[r] = new string[cols];
                    for (int c = 0; c < cols; c++) near[r][c] = _sheet.GetDisplay(startR + r, startC + c);
                }
                ctx.NearbyValues = near;
            }
            catch { }
            // Selection write policy + schema
            try
            {
                var policy = new SpreadsheetApp.Core.AI.SelectionWritePolicy();
                var writable = new System.Collections.Generic.List<int>();
                for (int dc = 0; dc < colsHint; dc++)
                {
                    int abs = sc + dc; if (abs >= 0 && abs < _sheet.Columns) writable.Add(abs);
                }
                policy.WritableColumns = writable.ToArray();
                int inputCol = sc > 0 ? sc - 1 : sc;
                if (inputCol >= 0 && inputCol < _sheet.Columns)
                {
                    policy.InputColumnIndex = inputCol;
                    bool inputInsideSelection = (inputCol >= sc && inputCol < sc + colsHint);
                    policy.AllowInputWritesForExistingRows = false;
                    policy.AllowInputWritesForEmptyRows = inputInsideSelection; // allow adding inputs only for empty rows when inside selection
                }
                policy.HeaderRowReadOnly = true;
                ctx.WritePolicy = policy;

                var schemas = new System.Collections.Generic.List<SpreadsheetApp.Core.AI.ColumnSchema>();
                for (int dc = 0; dc < colsHint; dc++)
                {
                    int abs = sc + dc; if (abs < 0 || abs >= _sheet.Columns) continue;
                    var col = new SpreadsheetApp.Core.AI.ColumnSchema
                    {
                        ColumnIndex = abs,
                        ColumnLetter = SpreadsheetApp.Core.CellAddress.ColumnIndexToName(abs),
                        Name = sr > 0 ? (_sheet.GetDisplay(sr - 1, abs) ?? string.Empty) : string.Empty,
                        Type = "text",
                        AllowEmpty = true
                    };
                    schemas.Add(col);
                }
                if (schemas.Count > 0) ctx.Schema = schemas.ToArray();
            }
            catch { }
            try
            {
                // Workbook summary
                int sheetCount = _sheets.Count;
                var list = new System.Collections.Generic.List<SpreadsheetApp.Core.AI.SheetSummary>(sheetCount);
                for (int i = 0; i < sheetCount; i++)
                {
                    var sh = _sheets[i];
                    int usedR = 0, usedC = 0;
                    int minUsedR = int.MaxValue, minUsedC = int.MaxValue;
                    for (int r = 0; r < sh.Rows; r++)
                    {
                        for (int c = 0; c < sh.Columns; c++)
                        {
                            var raw = sh.GetRaw(r, c);
                            if (!string.IsNullOrWhiteSpace(raw))
                            {
                                if (r + 1 > usedR) usedR = r + 1;
                                if (c + 1 > usedC) usedC = c + 1;
                                if (r < minUsedR) minUsedR = r;
                                if (c < minUsedC) minUsedC = c;
                            }
                        }
                    }
                    if (minUsedR == int.MaxValue) { minUsedR = 0; }
                    if (minUsedC == int.MaxValue) { minUsedC = 0; }

                    int hdrIdx = -1;
                    string[]? header = null;
                    if (usedR > 0 && usedC > 0)
                    {
                        hdrIdx = DetectHeaderRow(sh, usedR, usedC);
                        if (hdrIdx < 0) hdrIdx = minUsedR; // fall back to first used row
                        header = new string[usedC];
                        for (int c = 0; c < usedC; c++) header[c] = sh.GetDisplay(hdrIdx, c);
                    }

                    // Count data rows excluding header (rows with any non-empty cell below header up to usedR-1)
                    int dataCount = 0;
                    if (usedR > 0 && usedC > 0)
                    {
                        for (int r = hdrIdx + 1; r < usedR; r++)
                        {
                            bool any = false;
                            for (int c = 0; c < usedC; c++)
                            {
                                var raw = sh.GetRaw(r, c);
                                if (!string.IsNullOrWhiteSpace(raw)) { any = true; break; }
                            }
                            if (any) dataCount++;
                        }
                    }

                    string? usedTopLeft = null, usedBottomRight = null;
                    if (usedR > 0 && usedC > 0)
                    {
                        usedTopLeft = SpreadsheetApp.Core.CellAddress.ToAddress(minUsedR, minUsedC);
                        usedBottomRight = SpreadsheetApp.Core.CellAddress.ToAddress(usedR - 1, usedC - 1);
                    }

                    list.Add(new SpreadsheetApp.Core.AI.SheetSummary
                    {
                        Name = _sheetNames.Count > i ? _sheetNames[i] : $"Sheet{i + 1}",
                        UsedRows = usedR,
                        UsedCols = usedC,
                        HeaderRow = header,
                        HeaderRowIndex = hdrIdx,
                        DataRowCountExcludingHeader = dataCount,
                        UsedTopLeft = usedTopLeft,
                        UsedBottomRight = usedBottomRight
                    });
                }
                ctx.Workbook = list.ToArray();
            }
            catch { }

            return ctx;
        }

        // temp holder to thread selection values through ctx creation
        private string[][]? _lastSelVals;

        private void ApplyPlan(AIPlan plan)
        {
            // Log the AI action
            try { _aiActionLog.Record("(applied plan)", plan); } catch { }
            var edits = new List<(int row, int col, string? oldRaw, string? newRaw)>();
            var affected = new HashSet<(int r, int c)>();
            int? sheetAddedIndex = null; string? sheetAddedName = null;
            foreach (var cmd in plan.Commands)
            {
                if (cmd is SetValuesCommand set)
                {
                    // Handle append semantics: if StartRow < 0, append below the last non-empty in column
                    int baseRow = set.StartRow < 0 ? FindFirstEmptyBelow(set.StartCol, grid.CurrentCell?.RowIndex ?? 0) : set.StartRow;
                    int rows = set.Values.Length;
                    for (int r = 0; r < rows; r++)
                    {
                        int cols = set.Values[r]?.Length ?? 0;
                        for (int c = 0; c < cols; c++)
                        {
                            int rr = baseRow + r;
                            int cc = set.StartCol + c;
                            if (rr < 0 || rr >= _sheet.Rows || cc < 0 || cc >= _sheet.Columns) continue;
                            string? oldRaw = _sheet.GetRaw(rr, cc);
                            string newRaw = set.Values[r][c] ?? string.Empty;
                            _sheet.SetRaw(rr, cc, newRaw);
                            edits.Add((rr, cc, oldRaw, newRaw));
                            foreach (var ac in _sheet.RecalculateDirty(rr, cc)) affected.Add(ac);
                        }
                    }
                }
                else if (cmd is SetFormulaCommand sf)
                {
                    int rows = sf.Formulas.Length;
                    int cols = rows > 0 ? sf.Formulas[0].Length : 0;
                    for (int r = 0; r < rows; r++)
                    {
                        for (int c = 0; c < cols; c++)
                        {
                            int rr = sf.StartRow + r;
                            int cc = sf.StartCol + c;
                            if (rr < 0 || rr >= _sheet.Rows || cc < 0 || cc >= _sheet.Columns) continue;
                            string? oldRaw = _sheet.GetRaw(rr, cc);
                            string formula = sf.Formulas[r][c] ?? string.Empty;
                            // Ensure formulas are written as raw text; allow missing leading '=' (we'll add it)
                            string newRaw = string.IsNullOrEmpty(formula) ? string.Empty : (formula.StartsWith("=", StringComparison.Ordinal) ? formula : ("=" + formula));
                            _sheet.SetRaw(rr, cc, newRaw);
                            edits.Add((rr, cc, oldRaw, newRaw));
                            foreach (var ac in _sheet.RecalculateDirty(rr, cc)) affected.Add(ac);
                        }
                    }
                }
                else if (cmd is ClearRangeCommand cr)
                {
                    int rows = Math.Max(1, cr.Rows);
                    int cols = Math.Max(1, cr.Cols);
                    for (int r = 0; r < rows; r++)
                    {
                        for (int c = 0; c < cols; c++)
                        {
                            int rr = cr.StartRow + r;
                            int cc = cr.StartCol + c;
                            if (rr < 0 || rr >= _sheet.Rows || cc < 0 || cc >= _sheet.Columns) continue;
                            string? oldRaw = _sheet.GetRaw(rr, cc);
                            if (!string.IsNullOrEmpty(oldRaw))
                            {
                                _sheet.SetRaw(rr, cc, string.Empty);
                                edits.Add((rr, cc, oldRaw, string.Empty));
                                foreach (var ac in _sheet.RecalculateDirty(rr, cc)) affected.Add(ac);
                            }
                        }
                    }
                }
                else if (cmd is RenameSheetCommand rn)
                {
                    int targetIndex = _activeSheetIndex;
                    if (rn.Index1.HasValue)
                    {
                        targetIndex = Math.Max(0, Math.Min(_sheets.Count - 1, rn.Index1.Value - 1));
                    }
                    else if (!string.IsNullOrWhiteSpace(rn.OldName))
                    {
                        int idx = _sheetNames.FindIndex(n => string.Equals(n, rn.OldName, StringComparison.OrdinalIgnoreCase));
                        if (idx >= 0) targetIndex = idx;
                    }
                    string newName = string.IsNullOrWhiteSpace(rn.NewName) ? _sheetNames[targetIndex] : rn.NewName.Trim();
                    _sheetNames[targetIndex] = newName;
                    _activeSheetIndex = targetIndex;
                    RefreshTabs();
                }
                else if (cmd is SetTitleCommand st)
                {
                    for (int r = 0; r < st.Rows; r++)
                    {
                        for (int c = 0; c < st.Cols; c++)
                        {
                            int rr = st.StartRow + r;
                            int cc = st.StartCol + c;
                            if (rr < 0 || rr >= _sheet.Rows || cc < 0 || cc >= _sheet.Columns) continue;
                            string? oldRaw = _sheet.GetRaw(rr, cc);
                            string newRaw = st.Text;
                            _sheet.SetRaw(rr, cc, newRaw);
                            edits.Add((rr, cc, oldRaw, newRaw));
                            foreach (var ac in _sheet.RecalculateDirty(rr, cc)) affected.Add(ac);
                        }
                    }
                }
                else if (cmd is CreateSheetCommand cs)
                {
                    AddSheet(cs.Name);
                    sheetAddedIndex = _activeSheetIndex;
                    sheetAddedName = cs.Name;
                }
                else if (cmd is SortRangeCommand sr)
                {
                    int rows = Math.Max(1, sr.Rows);
                    int cols = Math.Max(1, sr.Cols);
                    int r0 = sr.StartRow;
                    int c0 = sr.StartCol;
                    int sortCol = Math.Max(0, Math.Min(_sheet.Columns - 1, sr.SortCol));
                    int rStart = r0;
                    int rCount = rows;
                    if (sr.HasHeader && rCount > 1)
                    {
                        // Keep header at first row
                        rStart = r0 + 1;
                        rCount = rows - 1;
                    }
                    // Snapshot current block raw values
                    var block = new string[rows][];
                    for (int r = 0; r < rows; r++)
                    {
                        block[r] = new string[cols];
                        for (int c = 0; c < cols; c++)
                        {
                            int rr = r0 + r; int cc = c0 + c;
                            block[r][c] = _sheet.GetRaw(rr, cc) ?? string.Empty;
                        }
                    }
                    // Build sortable list of row indices within the sortable segment
                    var idx = new System.Collections.Generic.List<int>();
                    for (int i = 0; i < rCount; i++) idx.Add(i);
                    // Comparison using evaluated value of sort column when numeric, else case-insensitive text
                    int sortRelCol = sortCol - c0; // may be outside block; clamp
                    if (sortRelCol < 0 || sortRelCol >= cols)
                    {
                        sortRelCol = 0; // default to first column inside region
                    }
                    idx.Sort((a, b) =>
                    {
                        int ra = rStart - r0 + a; // relative row inside block
                        int rb = rStart - r0 + b;
                        var va = _sheet.GetValue(r0 + ra, c0 + sortRelCol);
                        var vb = _sheet.GetValue(r0 + rb, c0 + sortRelCol);
                        int cmp;
                        if (va.Error == null && vb.Error == null && va.Number is double na && vb.Number is double nb)
                        {
                            cmp = na.CompareTo(nb);
                        }
                        else
                        {
                            var sa = va.ToDisplay();
                            var sb = vb.ToDisplay();
                            // empty last
                            bool ea = string.IsNullOrWhiteSpace(sa);
                            bool eb = string.IsNullOrWhiteSpace(sb);
                            if (ea && eb) cmp = 0; else if (ea) cmp = 1; else if (eb) cmp = -1; else cmp = string.Compare(sa, sb, StringComparison.OrdinalIgnoreCase);
                        }
                        if (string.Equals(sr.Order, "desc", StringComparison.OrdinalIgnoreCase)) cmp = -cmp;
                        return cmp;
                    });

                    // Construct new block with sorted rows (preserving header if present)
                    var sortedBlock = new string[rows][];
                    for (int r = 0; r < rows; r++) sortedBlock[r] = new string[cols];
                    if (sr.HasHeader && rows > 0)
                    {
                        Array.Copy(block[0], sortedBlock[0], cols);
                    }
                    for (int i = 0; i < rCount; i++)
                    {
                        int srcRel = rStart - r0 + idx[i];
                        int dstRel = rStart - r0 + i;
                        Array.Copy(block[srcRel], sortedBlock[dstRel], cols);
                    }
                    // Apply sorted block
                    for (int r = 0; r < rows; r++)
                    {
                        for (int c = 0; c < cols; c++)
                        {
                            int rr = r0 + r; int cc = c0 + c;
                            string? oldRaw = _sheet.GetRaw(rr, cc);
                            string newRaw = sortedBlock[r][c] ?? string.Empty;
                            if (oldRaw != newRaw)
                            {
                                _sheet.SetRaw(rr, cc, newRaw);
                                edits.Add((rr, cc, oldRaw, newRaw));
                                foreach (var ac in _sheet.RecalculateDirty(rr, cc)) affected.Add(ac);
                            }
                        }
                    }
                }
                else if (cmd is InsertRowsCommand ir)
                {
                    // Shift data down by ir.Count rows starting at ir.At
                    int at = Math.Max(0, Math.Min(_sheet.Rows - 1, ir.At));
                    int count = Math.Max(1, Math.Min(_sheet.Rows - at, ir.Count));
                    // Shift from bottom up to avoid overwrites
                    for (int r = _sheet.Rows - 1; r >= at + count; r--)
                    {
                        int srcRow = r - count;
                        for (int c = 0; c < _sheet.Columns; c++)
                        {
                            string? srcRaw = _sheet.GetRaw(srcRow, c);
                            string? dstRaw = _sheet.GetRaw(r, c);
                            if (srcRaw != dstRaw)
                            {
                                _sheet.SetRaw(r, c, srcRaw);
                                edits.Add((r, c, dstRaw, srcRaw));
                            }
                            // Copy format too
                            var srcFmt = _sheet.GetFormat(srcRow, c);
                            _sheet.SetFormat(r, c, srcFmt);
                        }
                    }
                    // Clear the inserted rows
                    for (int r = at; r < at + count && r < _sheet.Rows; r++)
                    {
                        for (int c = 0; c < _sheet.Columns; c++)
                        {
                            string? old = _sheet.GetRaw(r, c);
                            if (!string.IsNullOrEmpty(old))
                            {
                                _sheet.SetRaw(r, c, string.Empty);
                                edits.Add((r, c, old, string.Empty));
                            }
                            _sheet.SetFormat(r, c, null);
                        }
                    }
                    _sheet.Recalculate();
                    for (int r = at; r < _sheet.Rows; r++)
                        for (int c = 0; c < _sheet.Columns; c++)
                            affected.Add((r, c));
                }
                else if (cmd is DeleteRowsCommand dr)
                {
                    // Shift data up by dr.Count rows starting at dr.At
                    int at = Math.Max(0, Math.Min(_sheet.Rows - 1, dr.At));
                    int count = Math.Max(1, Math.Min(_sheet.Rows - at, dr.Count));
                    // Record old values for undo
                    for (int r = at; r < at + count && r < _sheet.Rows; r++)
                    {
                        for (int c = 0; c < _sheet.Columns; c++)
                        {
                            string? old = _sheet.GetRaw(r, c);
                            if (!string.IsNullOrEmpty(old))
                                edits.Add((r, c, old, string.Empty));
                        }
                    }
                    // Shift rows up
                    for (int r = at; r < _sheet.Rows - count; r++)
                    {
                        int srcRow = r + count;
                        for (int c = 0; c < _sheet.Columns; c++)
                        {
                            string? srcRaw = _sheet.GetRaw(srcRow, c);
                            string? dstRaw = _sheet.GetRaw(r, c);
                            _sheet.SetRaw(r, c, srcRaw);
                            if (srcRaw != dstRaw)
                                edits.Add((r, c, dstRaw, srcRaw));
                            var srcFmt = _sheet.GetFormat(srcRow, c);
                            _sheet.SetFormat(r, c, srcFmt);
                        }
                    }
                    // Clear the vacated rows at the bottom
                    for (int r = _sheet.Rows - count; r < _sheet.Rows; r++)
                    {
                        for (int c = 0; c < _sheet.Columns; c++)
                        {
                            string? old = _sheet.GetRaw(r, c);
                            if (!string.IsNullOrEmpty(old))
                            {
                                _sheet.SetRaw(r, c, string.Empty);
                                edits.Add((r, c, old, string.Empty));
                            }
                            _sheet.SetFormat(r, c, null);
                        }
                    }
                    _sheet.Recalculate();
                    for (int r = at; r < _sheet.Rows; r++)
                        for (int c = 0; c < _sheet.Columns; c++)
                            affected.Add((r, c));
                }
            }
            if (edits.Count > 0 || sheetAddedIndex.HasValue)
            {
                // Record as a single composite undo step
                var be = new List<UndoManager.BulkEdit>(edits.Count);
                foreach (var e in edits) be.Add(new UndoManager.BulkEdit(e.row, e.col, e.oldRaw, e.newRaw));
                (int index, string name)? sheetAddArg = sheetAddedIndex.HasValue && sheetAddedName != null
                    ? (sheetAddedIndex.Value, sheetAddedName)
                    : null;
                _undo.RecordComposite(be, sheetAddArg);
                try { RefreshDirtyOrFull(affected); } catch { }
                try { UpdateAiMenuItemsState(); } catch { }
            }

            // Post-apply error feedback loop: prompt to fix cells that evaluated to errors
            try
            {
                var seen = new HashSet<(int r, int c)>();
                var errs = new List<(int r, int c, string raw, string err)>();
                foreach (var e in edits)
                {
                    if (!seen.Add((e.row, e.col))) continue;
                    var val = _sheet.GetValue(e.row, e.col);
                    if (val.Error != null)
                    {
                        var raw = _sheet.GetRaw(e.row, e.col) ?? string.Empty;
                        errs.Add((e.row, e.col, raw, val.Error));
                    }
                }
                if (errs.Count > 0)
                {
                    var preview = new System.Text.StringBuilder();
                    int max = Math.Min(6, errs.Count);
                    for (int i = 0; i < max; i++)
                    {
                        var it = errs[i];
                        preview.Append(Core.CellAddress.ToAddress(it.r, it.c)).Append(" ").Append(it.raw).Append(" => #ERR: ").Append(it.err).AppendLine();
                    }
                    var resp = MessageBox.Show(this, $"Some cells errored. Attempt a fix?\n\n" + preview.ToString(), "AI", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (resp == DialogResult.Yes)
                    {
                        string prompt = BuildErrorRepairPrompt(errs);
                        if (_chatDockHost == null || _chatDockSplitter == null) { try { CreateDockedChatPane(); } catch { } }
                        if (_chatDockHost != null)
                        {
                            _chatDockHost.Visible = true;
                            if (_chatDockSplitter != null) _chatDockSplitter.Visible = true;
                            _chatPane?.SetPrompt(prompt, autoPlan: true);
                        }
                    }
                }
            }
            catch { }
        }

        private string BuildErrorRepairPrompt(System.Collections.Generic.List<(int r, int c, string raw, string err)> errs)
        {
            var sb = new System.Text.StringBuilder();
            sb.Append("Some cells produced errors after the last changes. Please propose fixes without adding unrelated changes. Errors: ");
            int limit = Math.Min(20, errs.Count);
            for (int i = 0; i < limit; i++)
            {
                var e = errs[i];
                sb.Append("[").Append(Core.CellAddress.ToAddress(e.r, e.c)).Append(": ")
                  .Append(e.raw).Append(" -> #ERR: ").Append(e.err).Append("] ");
            }
            return sb.ToString();
        }

        private int FindFirstEmptyBelow(int col, int fromRow)
        {
            int r = fromRow;
            while (r < _sheet.Rows)
            {
                var raw = _sheet.GetRaw(r, col);
                if (string.IsNullOrWhiteSpace(raw)) return r;
                r++;
            }
            return Math.Max(0, _sheet.Rows - 1);
        }

        private static int DetectHeaderRow(Core.Spreadsheet sh, int usedRows, int usedCols)
        {
            // Heuristic: first non-empty row with predominantly text values
            for (int r = 0; r < usedRows; r++)
            {
                int nonEmpty = 0, textCount = 0, numberCount = 0;
                for (int c = 0; c < usedCols; c++)
                {
                    var disp = sh.GetDisplay(r, c) ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(disp)) continue;
                    nonEmpty++;
                    if (double.TryParse(disp, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out _)) numberCount++;
                    else if (!disp.StartsWith("#ERR", StringComparison.Ordinal)) textCount++;
                }
                if (nonEmpty == 0) continue;
                if (textCount >= Math.Max(1, usedCols / 2) && textCount >= numberCount) return r;
            }
            return -1;
        }

        private void OpenAiSettings()
        {
            using var dlg = new UI.AI.SettingsDialog(_settings);
            if (dlg.ShowDialog(this) == DialogResult.OK)
            {
                _settings.AiEnabled = dlg.AiEnabled;
                _settings.Provider = dlg.Provider;
                _settings.Save();
                ApplySettings(_settings);
                SyncAiMenuState();
                if (!_settings.AiEnabled) { HideGhostSuggestions(); CancelInlineRequest(); }
            }
        }

        private void ApplySettings(SpreadsheetApp.Core.AppSettings s)
        {
            _settings = s;
            _aiInlineEnabled = s.AiEnabled;
            // Provider selection: Auto -> env detection
            string provider = s.Provider?.Trim() ?? "Auto";
            if (string.Equals(provider, "Auto", StringComparison.OrdinalIgnoreCase))
            {
                var oai = Environment.GetEnvironmentVariable("OPENAI_API_KEY");
                var claude = Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY");
                if (!string.IsNullOrWhiteSpace(oai)) provider = "OpenAI";
                else if (!string.IsNullOrWhiteSpace(claude)) provider = "Anthropic";
                else provider = "Mock";
            }
            switch (provider)
            {
                case "OpenAI":
                    {
                        var key = Environment.GetEnvironmentVariable("OPENAI_API_KEY");
                        _aiProvider = !string.IsNullOrWhiteSpace(key) ? new OpenAIProvider(key!) : new MockInferenceProvider();
                        _chatPlanner = !string.IsNullOrWhiteSpace(key) ? new ProviderChatPlanner("OpenAI") : new MockChatPlanner();
                        break;
                    }
                case "Anthropic":
                    {
                        var key = Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY");
                        _aiProvider = !string.IsNullOrWhiteSpace(key) ? new AnthropicProvider(key!) : new MockInferenceProvider();
                        _chatPlanner = !string.IsNullOrWhiteSpace(key) ? new ProviderChatPlanner("Anthropic") : new MockChatPlanner();
                        break;
                    }
                case "External":
                    {
                        if (!string.IsNullOrWhiteSpace(s.ExternalApiBaseUrl) && !string.IsNullOrWhiteSpace(s.GetApiKey()))
                            _aiProvider = new ExternalApiProvider(s.ExternalApiBaseUrl!, s.GetApiKey()!);
                        else
                            _aiProvider = new MockInferenceProvider();
                        _chatPlanner = new MockChatPlanner();
                        break;
                    }
                default:
                    _aiProvider = new MockInferenceProvider();
                    _chatPlanner = new MockChatPlanner();
                    break;
            }
        }

        private void SyncAiMenuState()
        {
            try
            {
                aiEnableInlineToolStripMenuItem.Checked = _aiInlineEnabled && _settings.AiEnabled;
                UpdateAiMenuItemsState();
            }
            catch { }
        }

        private bool ProviderReady()
        {
            if (!_settings.AiEnabled) return false;
            string provider = _settings.Provider?.Trim() ?? "Auto";
            if (provider == "Auto")
            {
                if (!string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("OPENAI_API_KEY"))) return true;
                if (!string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY"))) return true;
                return true; // falls back to Mock
            }
            if (provider == "Mock") return true;
            if (provider == "OpenAI") return !string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("OPENAI_API_KEY"));
            if (provider == "Anthropic") return !string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY"));
            if (provider == "External") return !string.IsNullOrWhiteSpace(_settings.ExternalApiBaseUrl) && !string.IsNullOrWhiteSpace(_settings.GetApiKey());
            return false;
        }

        // --- AI: Generate Fill ---
        private void OpenGenerateFill()
        {
            if (!_settings.AiEnabled) { MessageBox.Show(this, "AI is disabled in Settings.", "AI", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            if (grid.CurrentCell == null) { MessageBox.Show(this, "Select a start cell first.", "AI", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            // Determine selection shape; fall back to single cell with default rows/cols
            int startR = grid.CurrentCell.RowIndex;
            int startC = grid.CurrentCell.ColumnIndex;
            int selRows = 10;
            int selCols = 1;
            try
            {
                if (grid.SelectedCells != null && grid.SelectedCells.Count > 1)
                {
                    int minR = int.MaxValue, maxR = -1, minC = int.MaxValue, maxC = -1;
                    foreach (DataGridViewCell cell in grid.SelectedCells)
                    {
                        if (cell == null) continue;
                        if (cell.RowIndex < minR) minR = cell.RowIndex;
                        if (cell.RowIndex > maxR) maxR = cell.RowIndex;
                        if (cell.ColumnIndex < minC) minC = cell.ColumnIndex;
                        if (cell.ColumnIndex > maxC) maxC = cell.ColumnIndex;
                    }
                    if (minR <= maxR && minC <= maxC)
                    {
                        startR = minR;
                        startC = minC;
                        selRows = (maxR - minR + 1);
                        selCols = (maxC - minC + 1);
                    }
                }
            }
            catch { }
            string sheetName = _sheetNames.Count > _activeSheetIndex ? _sheetNames[_activeSheetIndex] : "Sheet";
            string? title = null;
            if (startR > 0)
            {
                var above = _sheet.GetRaw(startR - 1, startC);
                if (!string.IsNullOrWhiteSpace(above)) title = above;
            }
            var ctx = new AIContext { SheetName = sheetName, StartRow = startR, StartCol = startC, Rows = selRows, Cols = selCols, Title = title, Prompt = string.Empty };

            using var dlg = new GenerateFillDialog(_aiProvider, ctx);
            if (dlg.ShowDialog(this) == DialogResult.OK)
            {
                var cells = dlg.ResultCells;
                if (cells.Length == 0) return;
                ApplyCellsWithUndo(startR, startC, cells);
            }
        }

        private void ApplyCellsWithUndo(int startR, int startC, string[][] cells)
        {
            int rows = cells.Length;
            int cols = cells[0].Length;
            var edits = new List<(int row, int col, string? oldRaw, string? newRaw)>();
            var affected = new HashSet<(int r, int c)>();
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    int rr = startR + r;
                    int cc = startC + c;
                    if (rr < 0 || rr >= _sheet.Rows || cc < 0 || cc >= _sheet.Columns) continue;
                    string? oldRaw = _sheet.GetRaw(rr, cc);
                    string newRaw = cells[r][c] ?? string.Empty;
                    _sheet.SetRaw(rr, cc, newRaw);
                    edits.Add((rr, cc, oldRaw, newRaw));
                    foreach (var ac in _sheet.RecalculateDirty(rr, cc)) affected.Add(ac);
                }
            }
            if (edits.Count > 0)
            {
                _undo.RecordBulk(edits);
                try
                {
                    RefreshDirtyOrFull(affected);
                }
                catch { }
                try { UpdateAiMenuItemsState(); } catch { }
            }
        }

        private void NewSheet()
        {
            if (ConfirmDiscard())
            {
                _sheets.Clear(); _sheetNames.Clear(); _undos.Clear();
                InitializeSheet(new Spreadsheet(DefaultRows, DefaultCols));
            }
        }

        private bool ConfirmDiscard()
        {
            var res = MessageBox.Show(this, "Discard current sheet?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            return res == DialogResult.Yes;
        }

        private void SaveSheet()
        {
            // Legacy sync path retained; prefer SaveSheetAsync for UI
            _ = SaveSheetAsync();
        }

        private void OpenSheet()
        {
            // Legacy sync path retained; prefer OpenSheetAsync for UI
            _ = OpenSheetAsync();
        }

        private void SetUiBusy(bool busy)
        {
            try
            {
                UseWaitCursor = busy;
                Cursor = busy ? Cursors.WaitCursor : Cursors.Default;
                menuStrip1.Enabled = !busy;
                grid.Enabled = !busy;
            }
            catch { }
        }

        private async System.Threading.Tasks.Task SaveSheetAsync()
        {
            using var sfd = new SaveFileDialog
            {
                Filter = "Spreadsheet JSON (*.sheet.json)|*.sheet.json|All files (*.*)|*.*",
                FileName = "sheet.sheet.json"
            };
            if (sfd.ShowDialog(this) == DialogResult.OK)
            {
                SetUiBusy(true);
                try
                {
                    await IO.SpreadsheetIO.SaveToFileAsync(_sheet, sfd.FileName).ConfigureAwait(true);
                    AddRecentFile(sfd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, $"Save failed: {ex.Message}", "Save", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally { SetUiBusy(false); }
            }
        }

        private async System.Threading.Tasks.Task OpenSheetAsync()
        {
            using var ofd = new OpenFileDialog
            {
                Filter = "Spreadsheet JSON (*.sheet.json)|*.sheet.json|All files (*.*)|*.*"
            };
            if (ofd.ShowDialog(this) == DialogResult.OK)
            {
                SetUiBusy(true);
                try
                {
                    var loaded = await IO.SpreadsheetIO.LoadFromFileAsync(ofd.FileName).ConfigureAwait(true);
                    _sheets.Clear(); _sheetNames.Clear(); _undos.Clear();
                    InitializeSheet(loaded);
                    AddRecentFile(ofd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, $"Open failed: {ex.Message}", "Open", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally { SetUiBusy(false); }
            }
        }

        private async System.Threading.Tasks.Task SaveWorkbookAsync()
        {
            using var sfd = new SaveFileDialog
            {
                Filter = "Workbook JSON (*.workbook.json)|*.workbook.json|All files (*.*)|*.*",
                FileName = "workbook.workbook.json"
            };
            if (sfd.ShowDialog(this) == DialogResult.OK)
            {
                SetUiBusy(true);
                try
                {
                    await IO.SpreadsheetIO.SaveWorkbookToFileAsync(_sheets, _sheetNames, sfd.FileName).ConfigureAwait(true);
                    AddRecentFile(sfd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, $"Save workbook failed: {ex.Message}", "Save Workbook", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally { SetUiBusy(false); }
            }
        }

        private async System.Threading.Tasks.Task OpenWorkbookAsync()
        {
            using var ofd = new OpenFileDialog
            {
                Filter = "Workbook JSON (*.workbook.json)|*.workbook.json|All files (*.*)|*.*"
            };
            if (ofd.ShowDialog(this) == DialogResult.OK)
            {
                SetUiBusy(true);
                try
                {
                    var wb = await IO.SpreadsheetIO.LoadWorkbookFromFileAsync(ofd.FileName).ConfigureAwait(true);
                    _sheets.Clear(); _sheetNames.Clear(); _undos.Clear();
                    for (int i = 0; i < wb.Sheets.Count; i++)
                    {
                        _sheets.Add(wb.Sheets[i]);
                        _sheetNames.Add(i < wb.Names.Count ? wb.Names[i] : $"Sheet{i + 1}");
                        _undos.Add(new UndoManager());
                    }
                    _activeSheetIndex = 0;
                    InitializeSheet(_sheets[0]);
                    AddRecentFile(ofd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, $"Open workbook failed: {ex.Message}", "Open Workbook", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally { SetUiBusy(false); }
            }
        }

        // --- Recent files ---
        private string RecentFilesPath()
        {
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SpreadsheetApp");
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, "recent.txt");
        }

        private void LoadRecentFiles()
        {
            var path = RecentFilesPath();
            if (File.Exists(path))
            {
                try
                {
                    var lines = File.ReadAllLines(path);
                    recentFilesToolStripMenuItem.DropDownItems.Clear();
                    foreach (var l in lines)
                    {
                        var p = l.Trim(); if (string.IsNullOrEmpty(p)) continue;
                        var item = new ToolStripMenuItem(p);
                        item.Click += (_, __) => { if (File.Exists(p)) { var s = IO.SpreadsheetIO.LoadFromFile(p); _sheets.Clear(); _sheetNames.Clear(); _undos.Clear(); InitializeSheet(s); AddRecentFile(p); } else { MessageBox.Show(this, "File not found."); } };
                        recentFilesToolStripMenuItem.DropDownItems.Add(item);
                    }
                }
                catch { }
            }
        }

        private void SaveRecentFilesMenu(IEnumerable<string> files)
        {
            var path = RecentFilesPath();
            try { File.WriteAllLines(path, files); } catch { }
            LoadRecentFiles();
        }

        private void AddRecentFile(string path)
        {
            var items = new List<string>();
            foreach (ToolStripItem it in recentFilesToolStripMenuItem.DropDownItems)
            {
                if (it is ToolStripMenuItem mi && !string.IsNullOrWhiteSpace(mi.Text))
                    items.Add(mi.Text);
            }
            items.RemoveAll(p => string.Equals(p, path, StringComparison.OrdinalIgnoreCase));
            items.Insert(0, path);
            if (items.Count > 5) items.RemoveRange(5, items.Count - 5);
            SaveRecentFilesMenu(items);
        }

        private void RefreshTabs()
        {
            _suppressTabChange = true;
            try
            {
                tabs.TabPages.Clear();
                for (int i = 0; i < _sheets.Count; i++)
                {
                    var tp = new TabPage(_sheetNames[i]);
                    tabs.TabPages.Add(tp);
                }
                if (_activeSheetIndex >= 0 && _activeSheetIndex < tabs.TabPages.Count)
                    tabs.SelectedIndex = _activeSheetIndex;
            }
            finally
            {
                _suppressTabChange = false;
            }
        }

        private void Tabs_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (_suppressTabChange) return;
            if (tabs.SelectedIndex < 0 || tabs.SelectedIndex >= _sheets.Count) return;
            _activeSheetIndex = tabs.SelectedIndex;
            _sheet = _sheets[_activeSheetIndex];
            _undo = _undos[_activeSheetIndex];
            InitializeSheet(_sheet);
        }

        private void AddSheet()
        {
            string name = $"Sheet{_sheets.Count + 1}";
            _sheets.Add(new Spreadsheet(DefaultRows, DefaultCols));
            _sheetNames.Add(name);
            _undos.Add(new UndoManager());
            _activeSheetIndex = _sheets.Count - 1;
            _sheet = _sheets[_activeSheetIndex];
            _undo = _undos[_activeSheetIndex];
            InitializeSheet(_sheet);
        }

        private void AddSheet(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) name = $"Sheet{_sheets.Count + 1}";
            name = GetUniqueSheetName(name.Trim());
            _sheets.Add(new Spreadsheet(DefaultRows, DefaultCols));
            _sheetNames.Add(name);
            _undos.Add(new UndoManager());
            _activeSheetIndex = _sheets.Count - 1;
            _sheet = _sheets[_activeSheetIndex];
            _undo = _undos[_activeSheetIndex];
            InitializeSheet(_sheet);
        }

        private void AddSheetAt(int index, string name)
        {
            if (index < 0 || index > _sheets.Count) index = _sheets.Count;
            if (string.IsNullOrWhiteSpace(name)) name = $"Sheet{index + 1}";
            name = GetUniqueSheetName(name.Trim());
            _sheets.Insert(index, new Spreadsheet(DefaultRows, DefaultCols));
            _sheetNames.Insert(index, name);
            _undos.Insert(index, new UndoManager());
            _activeSheetIndex = index;
            _sheet = _sheets[_activeSheetIndex];
            _undo = _undos[_activeSheetIndex];
            InitializeSheet(_sheet);
        }

        private string GetUniqueSheetName(string baseName)
        {
            // Ensure no duplicate sheet names; auto-suffix " (2)", " (3)", ...
            string name = string.IsNullOrWhiteSpace(baseName) ? "Sheet" : baseName;
            if (!_sheetNames.Any(n => string.Equals(n, name, StringComparison.OrdinalIgnoreCase))) return name;
            int i = 2;
            while (true)
            {
                string candidate = $"{name} ({i})";
                if (!_sheetNames.Any(n => string.Equals(n, candidate, StringComparison.OrdinalIgnoreCase))) return candidate;
                i++;
                if (i > 1000) return Guid.NewGuid().ToString("N");
            }
        }

        private void RemoveSheetAt(int index)
        {
            if (_sheets.Count <= 1) { MessageBox.Show(this, "Cannot remove the only sheet."); return; }
            if (index < 0 || index >= _sheets.Count) return;
            _sheets.RemoveAt(index);
            _sheetNames.RemoveAt(index);
            _undos.RemoveAt(index);
            _activeSheetIndex = Math.Max(0, Math.Min(index, _sheets.Count - 1));
            _sheet = _sheets[_activeSheetIndex];
            _undo = _undos[_activeSheetIndex];
            InitializeSheet(_sheet);
        }

        private void RenameSheet()
        {
            if (_sheets.Count == 0) return;
            using var dlg = new Form { Text = "Rename Sheet", FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterParent, MinimizeBox = false, MaximizeBox = false, ClientSize = new System.Drawing.Size(300, 110) };
            var lbl = new Label { Text = "Name:", Left = 10, Top = 15, AutoSize = true };
            var tb = new TextBox { Left = 70, Top = 12, Width = 210, Text = _sheetNames[_activeSheetIndex] };
            var ok = new Button { Text = "OK", Left = 120, Top = 60, DialogResult = DialogResult.OK };
            var cancel = new Button { Text = "Cancel", Left = 200, Top = 60, DialogResult = DialogResult.Cancel };
            dlg.Controls.AddRange(new Control[] { lbl, tb, ok, cancel });
            dlg.AcceptButton = ok; dlg.CancelButton = cancel;
            if (dlg.ShowDialog(this) == DialogResult.OK)
            {
                _sheetNames[_activeSheetIndex] = string.IsNullOrWhiteSpace(tb.Text) ? _sheetNames[_activeSheetIndex] : tb.Text.Trim();
                RefreshTabs();
            }
        }

        private void RemoveSheet()
        {
            if (_sheets.Count <= 1) { MessageBox.Show(this, "Cannot remove the only sheet."); return; }
            int idx = _activeSheetIndex;
            _sheets.RemoveAt(idx);
            _sheetNames.RemoveAt(idx);
            _undos.RemoveAt(idx);
            _activeSheetIndex = Math.Max(0, idx - 1);
            _sheet = _sheets[_activeSheetIndex];
            _undo = _undos[_activeSheetIndex];
            InitializeSheet(_sheet);
        }

        private void ExportCsv()
        {
            using var sfd = new SaveFileDialog { Filter = "CSV (*.csv)|*.csv|All files (*.*)|*.*", FileName = "sheet.csv" };
            if (sfd.ShowDialog(this) == DialogResult.OK)
            {
                IO.SpreadsheetIO.ExportCsv(_sheet, sfd.FileName);
            }
        }

        private void ImportCsv()
        {
            using var ofd = new OpenFileDialog { Filter = "CSV (*.csv)|*.csv|All files (*.*)|*.*" };
            if (ofd.ShowDialog(this) == DialogResult.OK)
            {
                var imported = IO.SpreadsheetIO.ImportCsv(ofd.FileName);
                _sheets[_activeSheetIndex] = imported;
                _undos[_activeSheetIndex].Clear();
                _sheet = imported;
                InitializeSheet(_sheet);
            }
        }

        private async System.Threading.Tasks.Task ExportCsvAsync()
        {
            using var sfd = new SaveFileDialog { Filter = "CSV (*.csv)|*.csv|All files (*.*)|*.*", FileName = "sheet.csv" };
            if (sfd.ShowDialog(this) == DialogResult.OK)
            {
                SetUiBusy(true);
                try { await IO.SpreadsheetIO.ExportCsvAsync(_sheet, sfd.FileName).ConfigureAwait(true); }
                catch (Exception ex) { MessageBox.Show(this, $"Export failed: {ex.Message}", "CSV Export", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                finally { SetUiBusy(false); }
            }
        }

        private async System.Threading.Tasks.Task ImportCsvAsync()
        {
            using var ofd = new OpenFileDialog { Filter = "CSV (*.csv)|*.csv|All files (*.*)|*.*" };
            if (ofd.ShowDialog(this) == DialogResult.OK)
            {
                SetUiBusy(true);
                try
                {
                    var imported = await IO.SpreadsheetIO.ImportCsvAsync(ofd.FileName).ConfigureAwait(true);
                    _sheets[_activeSheetIndex] = imported;
                    _undos[_activeSheetIndex].Clear();
                    _sheet = imported;
                    InitializeSheet(_sheet);
                }
                catch (Exception ex) { MessageBox.Show(this, $"Import failed: {ex.Message}", "CSV Import", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                finally { SetUiBusy(false); }
            }
        }

        // Format menu handlers
        private void ToggleBoldFormat()
        {
            if (grid.CurrentCell == null) return;
            int r = grid.CurrentCell.RowIndex, c = grid.CurrentCell.ColumnIndex;
            var fmt = _sheet.GetFormat(r, c) ?? new CellFormat();
            fmt.Bold = !fmt.Bold;
            _sheet.SetFormat(r, c, fmt);
            RefreshGridDisplays();
        }

        private void ChooseTextColor()
        {
            if (grid.CurrentCell == null) return;
            using var cd = new ColorDialog();
            if (cd.ShowDialog(this) == DialogResult.OK)
            {
                int r = grid.CurrentCell.RowIndex, c = grid.CurrentCell.ColumnIndex;
                var fmt = _sheet.GetFormat(r, c) ?? new CellFormat();
                fmt.ForeColorArgb = cd.Color.ToArgb();
                _sheet.SetFormat(r, c, fmt);
                RefreshGridDisplays();
            }
        }

        private void ChooseFillColor()
        {
            if (grid.CurrentCell == null) return;
            using var cd = new ColorDialog();
            if (cd.ShowDialog(this) == DialogResult.OK)
            {
                int r = grid.CurrentCell.RowIndex, c = grid.CurrentCell.ColumnIndex;
                var fmt = _sheet.GetFormat(r, c) ?? new CellFormat();
                fmt.BackColorArgb = cd.Color.ToArgb();
                _sheet.SetFormat(r, c, fmt);
                RefreshGridDisplays();
            }
        }

        private void SetAlignment(string which)
        {
            if (grid.CurrentCell == null) return;
            int r = grid.CurrentCell.RowIndex, c = grid.CurrentCell.ColumnIndex;
            var fmt = _sheet.GetFormat(r, c) ?? new CellFormat();
            fmt.HAlign = which switch
            {
                "Center" => CellHAlign.Center,
                "Right" => CellHAlign.Right,
                _ => CellHAlign.Left
            };
            _sheet.SetFormat(r, c, fmt);
            RefreshGridDisplays();
        }

        private void SetNumberFormat(string fmtStr)
        {
            if (grid.CurrentCell == null) return;
            int r = grid.CurrentCell.RowIndex, c = grid.CurrentCell.ColumnIndex;
            var fmt = _sheet.GetFormat(r, c) ?? new CellFormat();
            fmt.NumberFormat = fmtStr == "General" ? null : fmtStr;
            _sheet.SetFormat(r, c, fmt);
            RefreshGridDisplays();
        }

        // --- Test Runner ---

        public void LoadWorkbookFromPath(string path)
        {
            var wb = IO.SpreadsheetIO.LoadWorkbookFromFile(path);
            _sheets.Clear(); _sheetNames.Clear(); _undos.Clear();
            for (int i = 0; i < wb.Sheets.Count; i++)
            {
                _sheets.Add(wb.Sheets[i]);
                _sheetNames.Add(i < wb.Names.Count ? wb.Names[i] : $"Sheet{i + 1}");
                _undos.Add(new UndoManager());
            }
            _activeSheetIndex = 0;
            InitializeSheet(_sheets[0]);
            _automationHistory.Clear();
        }

        private void OpenTestRunner()
        {
            using var runner = new TestRunnerForm(LoadWorkbookFromPath, RunChatStepAsync, SaveWorkbookSnapshotTo, ActivateSheetByName, ClearAutomationChatHistory, CaptureActiveSheetMap);
            runner.ShowDialog(this);
        }

        // --- Programmatic AI helpers for Test Runner ---
        private readonly System.Collections.Generic.List<SpreadsheetApp.Core.AI.ChatMessage> _automationHistory = new();

        public void ClearAutomationChatHistory()
        {
            _automationHistory.Clear();
        }

        public void SaveWorkbookSnapshotTo(string path)
        {
            try
            {
                IO.SpreadsheetIO.SaveWorkbookToFile(_sheets, _sheetNames, path);
            }
            catch { }
        }

        public System.Collections.Generic.Dictionary<string, string> CaptureActiveSheetMap()
        {
            var map = new System.Collections.Generic.Dictionary<string, string>(System.StringComparer.OrdinalIgnoreCase);
            try
            {
                var sh = _sheet;
                for (int r = 0; r < sh.Rows; r++)
                {
                    for (int c = 0; c < sh.Columns; c++)
                    {
                        var raw = sh.GetRaw(r, c);
                        if (!string.IsNullOrWhiteSpace(raw))
                        {
                            string addr = SpreadsheetApp.Core.CellAddress.ToAddress(r, c);
                            map[addr] = raw!;
                        }
                    }
                }
            }
            catch { }
            return map;
        }

        public void ActivateSheetByName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return;
            int idx = _sheetNames.FindIndex(n => string.Equals(n, name, StringComparison.OrdinalIgnoreCase));
            if (idx >= 0 && idx < _sheets.Count)
            {
                try { tabs.SelectedIndex = idx; } catch { _activeSheetIndex = idx; InitializeSheet(_sheets[idx]); }
            }
        }

        public async System.Threading.Tasks.Task<SpreadsheetApp.Core.AI.AIPlan> RunChatStepAsync(string prompt, string? location, bool apply, System.Threading.CancellationToken ct)
        {
            try
            {
                // Adjust selection/cursor
                if (!string.IsNullOrWhiteSpace(location))
                {
                    SetSelectionFromLocation(location!);
                }
            }
            catch { }

            AIContext ctx;
            try { ctx = BuildPlannerContext(); }
            catch (Exception ex)
            {
                try { System.IO.File.AppendAllText("crash.log", $"[{DateTime.Now:o}] BuildPlannerContext failed: {ex}\n"); } catch { }
                return new SpreadsheetApp.Core.AI.AIPlan();
            }
            // Ensure planner context strictly reflects the provided location shape (if any)
            try
            {
                if (!string.IsNullOrWhiteSpace(location))
                {
                    string loc = location!.Trim();
                    if (loc.Contains(":", StringComparison.Ordinal))
                    {
                        var parts = loc.Split(':');
                        if (parts.Length == 2 && Core.CellAddress.TryParse(parts[0].Trim(), out int r1, out int c1) && Core.CellAddress.TryParse(parts[1].Trim(), out int r2, out int c2))
                        {
                            int rStart = Math.Max(0, Math.Min(r1, r2));
                            int rEnd = Math.Min(_sheet.Rows - 1, Math.Max(r1, r2));
                            int cStart = Math.Max(0, Math.Min(c1, c2));
                            int cEnd = Math.Min(_sheet.Columns - 1, Math.Max(c1, c2));
                            int rows = rEnd - rStart + 1;
                            int cols = cEnd - cStart + 1;
                            ctx.StartRow = rStart;
                            ctx.StartCol = cStart;
                            ctx.Rows = rows;
                            ctx.Cols = cols;
                            // Title cell directly above top-left
                            ctx.Title = rStart > 0 ? (_sheet.GetRaw(rStart - 1, cStart) ?? string.Empty) : string.Empty;
                            // Provide exact selection values snapshot
                            var sel = new string[rows][];
                            for (int r = 0; r < rows; r++)
                            {
                                sel[r] = new string[cols];
                                for (int c = 0; c < cols; c++) sel[r][c] = _sheet.GetDisplay(rStart + r, cStart + c);
                            }
                            ctx.SelectionValues = sel;

                            // Recompute WritePolicy and Schema to match the overridden selection shape
                            try
                            {
                                var policy = new SpreadsheetApp.Core.AI.SelectionWritePolicy();
                                var writable = new System.Collections.Generic.List<int>();
                                for (int dc = 0; dc < cols; dc++)
                                {
                                    int abs = cStart + dc; if (abs >= 0 && abs < _sheet.Columns) writable.Add(abs);
                                }
                                policy.WritableColumns = writable.ToArray();
                                int inputCol = cStart > 0 ? cStart - 1 : cStart;
                                if (inputCol >= 0 && inputCol < _sheet.Columns)
                                {
                                    policy.InputColumnIndex = inputCol;
                                    bool inputInsideSelection = (inputCol >= cStart && inputCol < cStart + cols);
                                    policy.AllowInputWritesForExistingRows = false;
                                    policy.AllowInputWritesForEmptyRows = inputInsideSelection;
                                }
                                policy.HeaderRowReadOnly = true;
                                ctx.WritePolicy = policy;

                                var schemas = new System.Collections.Generic.List<SpreadsheetApp.Core.AI.ColumnSchema>();
                                for (int dc = 0; dc < cols; dc++)
                                {
                                    int abs = cStart + dc; if (abs < 0 || abs >= _sheet.Columns) continue;
                                    var col = new SpreadsheetApp.Core.AI.ColumnSchema
                                    {
                                        ColumnIndex = abs,
                                        ColumnLetter = SpreadsheetApp.Core.CellAddress.ColumnIndexToName(abs),
                                        Name = rStart > 0 ? (_sheet.GetDisplay(rStart - 1, abs) ?? string.Empty) : string.Empty,
                                        Type = "text",
                                        AllowEmpty = true
                                    };
                                    schemas.Add(col);
                                }
                                ctx.Schema = schemas.Count > 0 ? schemas.ToArray() : null;
                            }
                            catch { }
                        }
                    }
                    else if (Core.CellAddress.TryParse(loc, out int rr, out int cc))
                    {
                        // Single-cell anchor: keep existing row/col hints from BuildPlannerContext but ensure start position
                        ctx.StartRow = Math.Max(0, Math.Min(_sheet.Rows - 1, rr));
                        ctx.StartCol = Math.Max(0, Math.Min(_sheet.Columns - 1, cc));
                        ctx.Title = ctx.StartRow > 0 ? (_sheet.GetRaw(ctx.StartRow - 1, ctx.StartCol) ?? string.Empty) : string.Empty;
                        // For single-cell anchors, also align WritePolicy/Schema to 1x1
                        try
                        {
                            ctx.Rows = Math.Max(1, ctx.Rows);
                            ctx.Cols = Math.Max(1, ctx.Cols);
                            var policy = new SpreadsheetApp.Core.AI.SelectionWritePolicy
                            {
                                WritableColumns = new[] { ctx.StartCol },
                                InputColumnIndex = ctx.StartCol > 0 ? ctx.StartCol - 1 : ctx.StartCol,
                                AllowInputWritesForExistingRows = false,
                                AllowInputWritesForEmptyRows = (ctx.StartCol == (ctx.StartCol > 0 ? ctx.StartCol - 1 : ctx.StartCol))
                            };
                            ctx.WritePolicy = policy;
                            ctx.Schema = new[] { new SpreadsheetApp.Core.AI.ColumnSchema { ColumnIndex = ctx.StartCol, ColumnLetter = SpreadsheetApp.Core.CellAddress.ColumnIndexToName(ctx.StartCol), Name = ctx.StartRow > 0 ? (_sheet.GetDisplay(ctx.StartRow - 1, ctx.StartCol) ?? string.Empty) : string.Empty, Type = "text", AllowEmpty = true } };
                        }
                        catch { }
                    }
                }
            }
            catch { }
            try
            {
                if (_automationHistory.Count > 0)
                    ctx.Conversation = new System.Collections.Generic.List<SpreadsheetApp.Core.AI.ChatMessage>(_automationHistory);
            }
            catch { }

            // De-bias Title hint when the prompt forbids titles or requires values-only
            try
            {
                string p = (prompt ?? string.Empty);
                bool valuesOnly = p.IndexOf("set_values only", StringComparison.OrdinalIgnoreCase) >= 0;
                bool noTitles = valuesOnly || p.IndexOf("do not add title", StringComparison.OrdinalIgnoreCase) >= 0 || p.IndexOf("do not add titles", StringComparison.OrdinalIgnoreCase) >= 0;
                if (noTitles)
                {
                    ctx.Title = string.Empty;
                }
            }
            catch { }

            // AllowedCommands (explicit gating) derived from prompt when present or via light heuristics
            try
            {
                string p = prompt ?? string.Empty;
                if (p.IndexOf("set_values only", StringComparison.OrdinalIgnoreCase) >= 0 || p.IndexOf("Use set_values only", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    ctx.AllowedCommands = new[] { "set_values" };
                }
                else
                {
                    // Heuristic: if the instruction clearly asks to fill/write simple values (numbers/text/lists)
                    // and does NOT mention structural/advanced ops, restrict to set_values to reduce accidental formulas.
                    string low = p.ToLowerInvariant();
                    bool mentionsFormula = low.Contains("formula") || low.Contains("=a") || low.Contains("=sum") || low.Contains("sum(") || low.Contains("average(") || low.Contains("total row");
                    bool mentionsStructural = low.Contains("sort") || low.Contains("rename") || low.Contains("create sheet") || low.Contains("clear ") || low.Contains("insert ") || low.Contains("delete ") || low.Contains("title");
                    bool simpleFillIntent = (low.Contains("fill") || low.Contains("write") || low.Contains("list") || low.Contains("pairs") || low.Contains("numbers") || low.Contains("values"));
                    if (simpleFillIntent && !mentionsFormula && !mentionsStructural)
                    {
                        ctx.AllowedCommands = new[] { "set_values" };
                    }
                }
            }
            catch { }

            SpreadsheetApp.Core.AI.AIPlan plan;
            try
            {
                plan = await _chatPlanner.PlanAsync(ctx, prompt ?? string.Empty, ct).ConfigureAwait(true);
            }
            catch (Exception ex)
            {
                try { System.IO.File.AppendAllText("crash.log", $"[{DateTime.Now:o}] PlanAsync failed: {ex}\n"); } catch { }
                return new SpreadsheetApp.Core.AI.AIPlan();
            }

            // Update automation conversation history (mirror ChatAssistantForm)
            try
            {
                var userMsg = new SpreadsheetApp.Core.AI.ChatMessage { Role = "user", Content = prompt ?? string.Empty };
                var asstSummary = string.Join("; ", plan.Commands.Select(c => c.Summarize()));
                var asstMsg = new SpreadsheetApp.Core.AI.ChatMessage { Role = "assistant", Content = asstSummary };
                _automationHistory.Add(userMsg);
                _automationHistory.Add(asstMsg);
                if (_automationHistory.Count > 10) _automationHistory.RemoveRange(0, _automationHistory.Count - 10);
            }
            catch { }

            // Enforce values-only constraints before sanitation/apply when explicitly requested
            try
            {
                string p = (prompt ?? string.Empty);
                bool valuesOnly = p.IndexOf("set_values only", StringComparison.OrdinalIgnoreCase) >= 0;
                bool noTitles = valuesOnly || p.IndexOf("do not add title", StringComparison.OrdinalIgnoreCase) >= 0 || p.IndexOf("do not add titles", StringComparison.OrdinalIgnoreCase) >= 0;
                if (valuesOnly)
                {
                    plan.Commands.RemoveAll(c => c is not SpreadsheetApp.Core.AI.SetValuesCommand);
                }
                else if (noTitles)
                {
                    plan.Commands.RemoveAll(c => c is SpreadsheetApp.Core.AI.SetTitleCommand);
                }
            }
            catch { }

            if (apply && plan.Commands.Count > 0)
            {
                // In automation, if a range was provided, sanitize plan to selection bounds
                try
                {
                    if (!string.IsNullOrWhiteSpace(location) && location!.Contains(":", StringComparison.Ordinal))
                    {
                        var parts = location.Split(':');
                        if (parts.Length == 2 && Core.CellAddress.TryParse(parts[0].Trim(), out int r1, out int c1) && Core.CellAddress.TryParse(parts[1].Trim(), out int r2, out int c2))
                        {
                            int rStart = Math.Max(0, Math.Min(r1, r2));
                            int rEnd = Math.Min(_sheet.Rows - 1, Math.Max(r1, r2));
                            int cStart = Math.Max(0, Math.Min(c1, c2));
                            int cEnd = Math.Min(_sheet.Columns - 1, Math.Max(c1, c2));
                            plan = SanitizePlanToBounds(plan, rStart, cStart, rEnd, cEnd);
                        }
                    }
                }
                catch { }
                ApplyPlan(plan);
            }
            return plan;
        }

        private SpreadsheetApp.Core.AI.AIPlan SanitizePlanToBounds(SpreadsheetApp.Core.AI.AIPlan plan, int rStart, int cStart, int rEnd, int cEnd)
        {
            var outPlan = new SpreadsheetApp.Core.AI.AIPlan();
            outPlan.RawJson = plan.RawJson; outPlan.RawUser = plan.RawUser; outPlan.RawSystem = plan.RawSystem;
            foreach (var cmd in plan.Commands)
            {
                switch (cmd)
                {
                    case SpreadsheetApp.Core.AI.SetValuesCommand sv:
                    {
                        int rows = sv.Values.Length;
                        int a1 = sv.StartRow; int b1 = sv.StartCol; int a2 = sv.StartRow + rows - 1;
                        // Determine max width across rows to compute right edge
                        int maxCols = 0;
                        for (int r = 0; r < rows; r++) { int len = sv.Values[r]?.Length ?? 0; if (len > maxCols) maxCols = len; }
                        int b2 = (maxCols > 0) ? (sv.StartCol + maxCols - 1) : (sv.StartCol - 1);
                        int rr1 = Math.Max(rStart, a1); int cc1 = Math.Max(cStart, b1);
                        int rr2 = Math.Min(rEnd, a2); int cc2 = Math.Min(cEnd, b2);
                        if (rr1 > rr2 || cc1 > cc2) break;
                        int newCols = cc2 - cc1 + 1;

                        var rowsOut = new System.Collections.Generic.List<string[]>();
                        for (int rr = rr1; rr <= rr2; rr++)
                        {
                            int srcRowIndex = rr - a1;
                            if (srcRowIndex < 0 || srcRowIndex >= rows) continue;
                            int rowLen = sv.Values[srcRowIndex]?.Length ?? 0;
                            int srcFirstCol = b1;
                            int srcLastCol = b1 + Math.Max(0, rowLen) - 1;
                            int cStartInt = Math.Max(cc1, srcFirstCol);
                            int cEndInt = Math.Min(cc2, srcLastCol);
                            int overlap = cEndInt - cStartInt + 1;
                            if (overlap <= 0) continue; // compact rows with no overlap

                            var line = new string[newCols];
                            for (int i = 0; i < newCols; i++) line[i] = string.Empty;
                            for (int c = 0; c < overlap; c++)
                            {
                                int destCol = (cStartInt - cc1) + c;
                                int srcCol = (cStartInt - srcFirstCol) + c;
                                line[destCol] = sv.Values[srcRowIndex][srcCol];
                            }
                            rowsOut.Add(line);
                        }
                        if (rowsOut.Count > 0)
                        {
                            // Header-echo suppression: drop first row if it matches header above the selection
                            try
                            {
                                if (rr1 > 0)
                                {
                                    int matches = 0, nonEmpty = 0;
                                    for (int c = 0; c < newCols; c++)
                                    {
                                        var hdr = _sheet.GetDisplay(rr1 - 1, cc1 + c) ?? string.Empty;
                                        var first = rowsOut[0][c] ?? string.Empty;
                                        if (!string.IsNullOrWhiteSpace(hdr))
                                        {
                                            nonEmpty++;
                                            if (string.Equals(hdr, first, StringComparison.OrdinalIgnoreCase)) matches++;
                                        }
                                    }
                                    if (nonEmpty >= 2 && matches >= nonEmpty)
                                    {
                                        rowsOut.RemoveAt(0);
                                    }
                                }
                            }
                            catch { }

                            if (rowsOut.Count > 0)
                            {
                                outPlan.Commands.Add(new SpreadsheetApp.Core.AI.SetValuesCommand { StartRow = rr1, StartCol = cc1, Values = rowsOut.ToArray() });
                            }
                        }
                        break;
                    }
                    case SpreadsheetApp.Core.AI.SetFormulaCommand sf:
                    {
                        int rows = sf.Formulas.Length; int cols = rows > 0 ? sf.Formulas[0].Length : 0;
                        int a1 = sf.StartRow; int b1 = sf.StartCol; int a2 = sf.StartRow + rows - 1; int b2 = sf.StartCol + cols - 1;
                        int rr1 = Math.Max(rStart, a1); int cc1 = Math.Max(cStart, b1);
                        int rr2 = Math.Min(rEnd, a2); int cc2 = Math.Min(cEnd, b2);
                        if (rr1 > rr2 || cc1 > cc2) break;
                        int newRows = rr2 - rr1 + 1; int newCols = cc2 - cc1 + 1;
                        var f = new string[newRows][];
                        for (int r = 0; r < newRows; r++)
                        {
                            f[r] = new string[newCols];
                            for (int c = 0; c < newCols; c++)
                            {
                                f[r][c] = sf.Formulas[(rr1 - a1) + r][(cc1 - b1) + c];
                            }
                        }
                        outPlan.Commands.Add(new SpreadsheetApp.Core.AI.SetFormulaCommand { StartRow = rr1, StartCol = cc1, Formulas = f });
                        break;
                    }
                    case SpreadsheetApp.Core.AI.ClearRangeCommand cr:
                    {
                        int a1 = cr.StartRow; int b1 = cr.StartCol; int a2 = cr.StartRow + Math.Max(1, cr.Rows) - 1; int b2 = cr.StartCol + Math.Max(1, cr.Cols) - 1;
                        int rr1 = Math.Max(rStart, a1); int cc1 = Math.Max(cStart, b1);
                        int rr2 = Math.Min(rEnd, a2); int cc2 = Math.Min(cEnd, b2);
                        if (rr1 > rr2 || cc1 > cc2) break;
                        outPlan.Commands.Add(new SpreadsheetApp.Core.AI.ClearRangeCommand { StartRow = rr1, StartCol = cc1, Rows = rr2 - rr1 + 1, Cols = cc2 - cc1 + 1 });
                        break;
                    }
                    case SpreadsheetApp.Core.AI.SortRangeCommand sr:
                    {
                        int a1 = sr.StartRow; int b1 = sr.StartCol; int a2 = sr.StartRow + Math.Max(1, sr.Rows) - 1; int b2 = sr.StartCol + Math.Max(1, sr.Cols) - 1;
                        int rr1 = Math.Max(rStart, a1); int cc1 = Math.Max(cStart, b1);
                        int rr2 = Math.Min(rEnd, a2); int cc2 = Math.Min(cEnd, b2);
                        if (rr1 > rr2 || cc1 > cc2) break;
                        var clone = new SpreadsheetApp.Core.AI.SortRangeCommand { StartRow = rr1, StartCol = cc1, Rows = rr2 - rr1 + 1, Cols = cc2 - cc1 + 1, Order = sr.Order, HasHeader = sr.HasHeader, SortCol = sr.SortCol };
                        outPlan.Commands.Add(clone);
                        break;
                    }
                    case SpreadsheetApp.Core.AI.SetTitleCommand st:
                    {
                        int a1 = st.StartRow; int b1 = st.StartCol; int a2 = st.StartRow + Math.Max(1, st.Rows) - 1; int b2 = st.StartCol + Math.Max(1, st.Cols) - 1;
                        int rr1 = Math.Max(rStart, a1); int cc1 = Math.Max(cStart, b1);
                        int rr2 = Math.Min(rEnd, a2); int cc2 = Math.Min(cEnd, b2);
                        if (rr1 > rr2 || cc1 > cc2) break;
                        outPlan.Commands.Add(new SpreadsheetApp.Core.AI.SetTitleCommand { StartRow = rr1, StartCol = cc1, Rows = rr2 - rr1 + 1, Cols = cc2 - cc1 + 1, Text = st.Text });
                        break;
                    }
                    default:
                        outPlan.Commands.Add(cmd);
                        break;
                }
            }
            return outPlan;
        }

        private void SetSelectionFromLocation(string loc)
        {
            if (string.IsNullOrWhiteSpace(loc) || grid.RowCount == 0 || grid.ColumnCount == 0) return;
            grid.ClearSelection();
            if (loc.Contains(":", StringComparison.Ordinal))
            {
                var parts = loc.Split(':');
                if (parts.Length == 2 && Core.CellAddress.TryParse(parts[0].Trim(), out int r1, out int c1) && Core.CellAddress.TryParse(parts[1].Trim(), out int r2, out int c2))
                {
                    int rStart = Math.Max(0, Math.Min(r1, r2));
                    int rEnd = Math.Min(_sheet.Rows - 1, Math.Max(r1, r2));
                    int cStart = Math.Max(0, Math.Min(c1, c2));
                    int cEnd = Math.Min(_sheet.Columns - 1, Math.Max(c1, c2));
                    for (int r = rStart; r <= rEnd; r++)
                    {
                        for (int c = cStart; c <= cEnd; c++)
                        {
                            try { grid[c, r].Selected = true; } catch { }
                        }
                    }
                    try
                    {
                        grid.CurrentCell = grid[cStart, rStart];
                        // Ensure the selected region is brought into view
                        try { grid.FirstDisplayedScrollingRowIndex = Math.Max(0, Math.Min(rStart, grid.RowCount - 1)); } catch { }
                        try { grid.FirstDisplayedScrollingColumnIndex = Math.Max(0, Math.Min(cStart, grid.ColumnCount - 1)); } catch { }
                    }
                    catch { }
                    return;
                }
            }
            // Single address
            if (Core.CellAddress.TryParse(loc.Trim(), out int rr, out int cc))
            {
                rr = Math.Max(0, Math.Min(_sheet.Rows - 1, rr));
                cc = Math.Max(0, Math.Min(_sheet.Columns - 1, cc));
                try
                {
                    grid.CurrentCell = grid[cc, rr]; grid[cc, rr].Selected = true;
                    try { grid.FirstDisplayedScrollingRowIndex = Math.Max(0, Math.Min(rr, grid.RowCount - 1)); } catch { }
                    try { grid.FirstDisplayedScrollingColumnIndex = Math.Max(0, Math.Min(cc, grid.ColumnCount - 1)); } catch { }
                }
                catch { }
            }
        }
    }
}
