using System;
using System.Globalization;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using SpreadsheetApp.Core;
using SpreadsheetApp.Core.AI;
using SpreadsheetApp.UI.AI;

namespace SpreadsheetApp.UI
{
    public partial class MainForm : Form
    {
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
        private readonly System.Collections.Generic.Dictionary<string, string[]> _aiInlineCache = new();
        private SpreadsheetApp.Core.AppSettings _settings = SpreadsheetApp.Core.AppSettings.Load();

        public MainForm()
        {
            InitializeComponent();
            try { LoadRecentFiles(); } catch { }
            // Inline AI setup
            _aiInlineTimer.Interval = 200;
            _aiInlineTimer.Tick += (_, __) => { _aiInlineTimer.Stop(); _ = TriggerInlineSuggestionAsync(); };
            CreateGhostUI();
            ApplySettings(_settings);
            SyncAiMenuState();
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
            if (keyData == (Keys.Control | Keys.Shift | Keys.C))
            {
                OpenChatAssistant();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
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
                RefreshGridValues();
                grid.Invalidate();
                grid.Refresh();
                UpdateStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Error during calculation: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            // Schedule inline suggestions after edit
            if (_aiInlineEnabled) ScheduleInlineSuggestion();
        }

        private void Undo()
        {
            if (_undo.TryUndoComposite(out var editsComp, out var sheetAddComp))
            {
                // Undo bulk edits first (set to old values), then undo sheet adds
                if (editsComp.Count > 0)
                {
                    _suppressRecord = true;
                    try
                    {
                        foreach (var e in editsComp)
                            _sheet.SetRaw(e.Row, e.Col, e.OldRaw);
                    }
                    finally { _suppressRecord = false; }
                }
                if (sheetAddComp.HasValue)
                {
                    RemoveSheetAt(sheetAddComp.Value.index);
                }
                try { RefreshGridValues(); } catch { }
            }
            else if (_undo.TryUndoSheetAdd(out int sheetIndex, out string sheetName))
            {
                RemoveSheetAt(sheetIndex);
            }
            else if (_undo.TryUndoBulk(out var edits))
            {
                _suppressRecord = true;
                try
                {
                    foreach (var e in edits)
                    {
                        _sheet.SetRaw(e.row, e.col, e.raw);
                    }
                }
                finally { _suppressRecord = false; }
                try { RefreshGridValues(); } catch { }
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
                try { RefreshGridValues(); } finally { _suppressRecord = false; }
                if (r >= 0 && r < grid.RowCount && c >= 0 && c < grid.ColumnCount)
                {
                    grid.CurrentCell = grid[c, r];
                }
            }
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
                    try
                    {
                        foreach (var e in editsComp)
                            _sheet.SetRaw(e.Row, e.Col, e.NewRaw);
                    }
                    finally { _suppressRecord = false; }
                }
                try { RefreshGridValues(); } catch { }
            }
            else if (_undo.TryRedoSheetAdd(out int sheetIndex, out string sheetName))
            {
                AddSheetAt(sheetIndex, sheetName);
            }
            else if (_undo.TryRedoBulk(out var edits))
            {
                _suppressRecord = true;
                try
                {
                    foreach (var e in edits)
                    {
                        _sheet.SetRaw(e.row, e.col, e.raw);
                    }
                }
                finally { _suppressRecord = false; }
                try { RefreshGridValues(); } catch { }
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
                try { RefreshGridValues(); } finally { _suppressRecord = false; }
                if (r >= 0 && r < grid.RowCount && c >= 0 && c < grid.ColumnCount)
                {
                    grid.CurrentCell = grid[c, r];
                }
            }
        }

        private void RefreshGridValues()
        {
            _sheet.Recalculate();
            RefreshGridDisplays();
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
                        if (fmt.NumberFormat == "0.00")
                            return d.ToString("0.00", CultureInfo.InvariantCulture);
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
        }

        private void UpdateStatus()
        {
            if (grid.CurrentCell == null)
            {
                statusCell.Text = "Cell: -";
                statusRaw.Text = "Raw: ";
                statusValue.Text = "Value: ";
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
        }

        // --- Copy / Paste / Cut ---
        private void CopyCell()
        {
            if (grid.CurrentCell == null) return;
            int r = grid.CurrentCell.RowIndex;
            int c = grid.CurrentCell.ColumnIndex;
            var raw = _sheet.GetRaw(r, c) ?? string.Empty;
            try { Clipboard.SetText(raw); } catch { /* ignore */ }
        }

        private void PasteCell()
        {
            if (grid.CurrentCell == null) return;
            string text = string.Empty;
            try { if (Clipboard.ContainsText()) text = Clipboard.GetText(); } catch { }
            int r = grid.CurrentCell.RowIndex;
            int c = grid.CurrentCell.ColumnIndex;
            var oldRaw = _sheet.GetRaw(r, c);
            _suppressRecord = true;
            try { _sheet.SetRaw(r, c, text); }
            finally { _suppressRecord = false; }
            _undo.RecordSet(r, c, oldRaw, text);
            try
            {
                RefreshGridValues();
                grid.Invalidate();
                grid.Refresh();
                UpdateStatus();
            }
            catch { }
        }

        private void CutCell()
        {
            if (grid.CurrentCell == null) return;
            int r = grid.CurrentCell.RowIndex;
            int c = grid.CurrentCell.ColumnIndex;
            var oldRaw = _sheet.GetRaw(r, c) ?? string.Empty;
            try { Clipboard.SetText(oldRaw); } catch { }
            _suppressRecord = true;
            try { _sheet.SetRaw(r, c, string.Empty); }
            finally { _suppressRecord = false; }
            _undo.RecordSet(r, c, oldRaw, string.Empty);
            try
            {
                RefreshGridValues();
                grid.Invalidate();
                grid.Refresh();
                UpdateStatus();
            }
            catch { }
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
                    }
                }
                if (edits.Count > 0)
                {
                    _undo.RecordBulk(edits);
                    RefreshGridValues();
                    grid.Invalidate();
                    grid.Refresh();
                    UpdateStatus();
                }
            }
            catch { }
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
                RefreshGridValues(); grid.Invalidate(); grid.Refresh(); UpdateStatus();

                // Move to next
                _findPosRow = r; _findPosCol = c;
                DoFindNext(_lastFind, _lastMatchCase);
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
                    }
                }
            }
            if (any)
            {
                RefreshGridValues(); grid.Invalidate(); grid.Refresh(); UpdateStatus();
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
                aiOpenChatToolStripMenuItem.Enabled = _settings.AiEnabled; // chat uses mock planner for now
                aiEnableInlineToolStripMenuItem.Enabled = _settings.AiEnabled && providerReady;
                if (!_settings.AiEnabled || !providerReady)
                {
                    aiEnableInlineToolStripMenuItem.Checked = false;
                }
            }
            catch { }
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
            Func<AIContext> getCtx = () =>
            {
                int sr = grid.CurrentCell?.RowIndex ?? 0;
                int sc = grid.CurrentCell?.ColumnIndex ?? 0;
                string? title = sr > 0 ? _sheet.GetRaw(sr - 1, sc) : null;
                return new AIContext
                {
                    SheetName = _sheetNames.Count > _activeSheetIndex ? _sheetNames[_activeSheetIndex] : "Sheet",
                    StartRow = sr,
                    StartCol = sc,
                    Rows = 5,
                    Cols = 1,
                    Title = title,
                    Prompt = string.Empty
                };
            };
            void apply(AIPlan plan) => ApplyPlan(plan);
            using var dlg = new ChatAssistantForm(_chatPlanner, getCtx, apply);
            dlg.ShowDialog(this);
        }

        private void ApplyPlan(AIPlan plan)
        {
            var edits = new List<(int row, int col, string? oldRaw, string? newRaw)>();
            int? sheetAddedIndex = null; string? sheetAddedName = null;
            foreach (var cmd in plan.Commands)
            {
                if (cmd is SetValuesCommand set)
                {
                    // Handle append semantics: if StartRow < 0, append below the last non-empty in column
                    int baseRow = set.StartRow < 0 ? FindFirstEmptyBelow(set.StartCol, grid.CurrentCell?.RowIndex ?? 0) : set.StartRow;
                    int rows = set.Values.Length;
                    int cols = rows > 0 ? set.Values[0].Length : 0;
                    for (int r = 0; r < rows; r++)
                    {
                        for (int c = 0; c < cols; c++)
                        {
                            int rr = baseRow + r;
                            int cc = set.StartCol + c;
                            if (rr < 0 || rr >= _sheet.Rows || cc < 0 || cc >= _sheet.Columns) continue;
                            string? oldRaw = _sheet.GetRaw(rr, cc);
                            string newRaw = set.Values[r][c] ?? string.Empty;
                            _sheet.SetRaw(rr, cc, newRaw);
                            edits.Add((rr, cc, oldRaw, newRaw));
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
                            }
                        }
                    }
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
                        }
                    }
                }
                else if (cmd is CreateSheetCommand cs)
                {
                    AddSheet(cs.Name);
                    sheetAddedIndex = _activeSheetIndex;
                    sheetAddedName = cs.Name;
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
                try { RefreshGridValues(); grid.Invalidate(); grid.Refresh(); UpdateStatus(); } catch { }
            }
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
                }
            }
            if (edits.Count > 0)
            {
                _undo.RecordBulk(edits);
                try
                {
                    RefreshGridValues(); grid.Invalidate(); grid.Refresh(); UpdateStatus();
                }
                catch { }
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
            _sheets.Add(new Spreadsheet(DefaultRows, DefaultCols));
            _sheetNames.Add(name.Trim());
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
            _sheets.Insert(index, new Spreadsheet(DefaultRows, DefaultCols));
            _sheetNames.Insert(index, name.Trim());
            _undos.Insert(index, new UndoManager());
            _activeSheetIndex = index;
            _sheet = _sheets[_activeSheetIndex];
            _undo = _undos[_activeSheetIndex];
            InitializeSheet(_sheet);
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
    }
}
