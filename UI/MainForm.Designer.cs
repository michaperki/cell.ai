using System.Drawing;
using System.Windows.Forms;

namespace SpreadsheetApp.UI
{
    partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;
        private MenuStrip menuStrip1 = null!;
        private StatusStrip statusStrip1 = null!;
        private ToolStripStatusLabel statusCell = null!;
        private ToolStripStatusLabel statusRaw = null!;
        private ToolStripStatusLabel statusValue = null!;
        private ToolStripMenuItem fileToolStripMenuItem = null!;
        private ToolStripMenuItem newToolStripMenuItem = null!;
        private ToolStripMenuItem openToolStripMenuItem = null!;
        private ToolStripMenuItem saveToolStripMenuItem = null!;
        private ToolStripMenuItem exitToolStripMenuItem = null!;
        private ToolStripMenuItem editToolStripMenuItem = null!;
        private ToolStripMenuItem formatToolStripMenuItem = null!;
        private ToolStripMenuItem aiToolStripMenuItem = null!;
        private ToolStripMenuItem aiGenerateFillToolStripMenuItem = null!;
        private ToolStripMenuItem aiEnableInlineToolStripMenuItem = null!;
        private ToolStripMenuItem aiAcceptInlineToolStripMenuItem = null!;
        private ToolStripMenuItem aiOpenChatToolStripMenuItem = null!;
        private ToolStripMenuItem aiToggleChatPaneToolStripMenuItem = null!;
        private ToolStripMenuItem aiSettingsToolStripMenuItem = null!;
        private ToolStripMenuItem undoToolStripMenuItem = null!;
        private ToolStripMenuItem redoToolStripMenuItem = null!;
        private ToolStripMenuItem copyToolStripMenuItem = null!;
        private ToolStripMenuItem pasteToolStripMenuItem = null!;
        private ToolStripMenuItem cutToolStripMenuItem = null!;
        private ToolStripMenuItem findToolStripMenuItem = null!;
        private ToolStripMenuItem replaceToolStripMenuItem = null!;
        private ToolStripMenuItem recalcToolStripMenuItem = null!;
        private ToolStripMenuItem clearContentsToolStripMenuItem = null!;
        private ToolStripMenuItem sheetsToolStripMenuItem = null!;
        private ToolStripMenuItem addSheetToolStripMenuItem = null!;
        private ToolStripMenuItem renameSheetToolStripMenuItem = null!;
        private ToolStripMenuItem removeSheetToolStripMenuItem = null!;
        private TabControl tabs = null!;
        private ToolStripMenuItem importCsvToolStripMenuItem = null!;
        private ToolStripMenuItem exportCsvToolStripMenuItem = null!;
        private ToolStripMenuItem recentFilesToolStripMenuItem = null!;
        private ToolStripMenuItem openWorkbookToolStripMenuItem = null!;
        private ToolStripMenuItem saveWorkbookToolStripMenuItem = null!;
        private ToolStripMenuItem testToolStripMenuItem = null!;
        private ToolStripMenuItem testRunnerToolStripMenuItem = null!;
        private DataGridView grid = null!;
        private Panel _formulaBarPanel = null!;
        private TextBox _cellNameBox = null!;
        private TextBox _formulaBar = null!;
        private ToolStripMenuItem helpToolStripMenuItem = null!;
        private ToolStripMenuItem docsViewerToolStripMenuItem = null!;
        private ToolStripMenuItem exportDocsJsonToolStripMenuItem = null!;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            menuStrip1 = new MenuStrip();
            statusStrip1 = new StatusStrip();
            tabs = new TabControl();
            fileToolStripMenuItem = new ToolStripMenuItem();
            newToolStripMenuItem = new ToolStripMenuItem();
            openToolStripMenuItem = new ToolStripMenuItem();
            saveToolStripMenuItem = new ToolStripMenuItem();
            importCsvToolStripMenuItem = new ToolStripMenuItem();
            exportCsvToolStripMenuItem = new ToolStripMenuItem();
            recentFilesToolStripMenuItem = new ToolStripMenuItem();
            openWorkbookToolStripMenuItem = new ToolStripMenuItem();
            saveWorkbookToolStripMenuItem = new ToolStripMenuItem();
            exitToolStripMenuItem = new ToolStripMenuItem();
            editToolStripMenuItem = new ToolStripMenuItem();
            formatToolStripMenuItem = new ToolStripMenuItem();
            undoToolStripMenuItem = new ToolStripMenuItem();
            redoToolStripMenuItem = new ToolStripMenuItem();
            copyToolStripMenuItem = new ToolStripMenuItem();
            pasteToolStripMenuItem = new ToolStripMenuItem();
            cutToolStripMenuItem = new ToolStripMenuItem();
            findToolStripMenuItem = new ToolStripMenuItem();
            replaceToolStripMenuItem = new ToolStripMenuItem();
            clearContentsToolStripMenuItem = new ToolStripMenuItem();
            recalcToolStripMenuItem = new ToolStripMenuItem();
            sheetsToolStripMenuItem = new ToolStripMenuItem();
            addSheetToolStripMenuItem = new ToolStripMenuItem();
            renameSheetToolStripMenuItem = new ToolStripMenuItem();
            removeSheetToolStripMenuItem = new ToolStripMenuItem();
            testToolStripMenuItem = new ToolStripMenuItem();
            testRunnerToolStripMenuItem = new ToolStripMenuItem();
            grid = new DataGridView();

            SuspendLayout();

            // menuStrip1
            menuStrip1.Items.AddRange(new ToolStripItem[] { fileToolStripMenuItem, editToolStripMenuItem, formatToolStripMenuItem, sheetsToolStripMenuItem });
            menuStrip1.Location = new System.Drawing.Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new System.Drawing.Size(1000, 24);
            menuStrip1.TabIndex = 0;
            menuStrip1.Text = "menuStrip1";

            // fileToolStripMenuItem
            fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            fileToolStripMenuItem.Text = "File";
            fileToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { newToolStripMenuItem, openToolStripMenuItem, saveToolStripMenuItem, new ToolStripSeparator(), openWorkbookToolStripMenuItem, saveWorkbookToolStripMenuItem, new ToolStripSeparator(), importCsvToolStripMenuItem, exportCsvToolStripMenuItem, new ToolStripSeparator(), recentFilesToolStripMenuItem, new ToolStripSeparator(), exitToolStripMenuItem });

            // newToolStripMenuItem
            newToolStripMenuItem.Name = "newToolStripMenuItem";
            newToolStripMenuItem.Text = "New";
            newToolStripMenuItem.ShortcutKeys = Keys.Control | Keys.N;
            newToolStripMenuItem.Click += (_, __) => NewSheet();

            // openToolStripMenuItem
            openToolStripMenuItem.Name = "openToolStripMenuItem";
            openToolStripMenuItem.Text = "Open...";
            openToolStripMenuItem.ShortcutKeys = Keys.Control | Keys.O;
            openToolStripMenuItem.Click += async (_, __) => await OpenSheetAsync();

            // saveToolStripMenuItem
            saveToolStripMenuItem.Name = "saveToolStripMenuItem";
            saveToolStripMenuItem.Text = "Save As...";
            saveToolStripMenuItem.ShortcutKeys = Keys.Control | Keys.S;
            saveToolStripMenuItem.Click += async (_, __) => await SaveSheetAsync();

            // importCsvToolStripMenuItem
            importCsvToolStripMenuItem.Name = "importCsvToolStripMenuItem";
            importCsvToolStripMenuItem.Text = "Import CSV...";
            importCsvToolStripMenuItem.Click += async (_, __) => await ImportCsvAsync();

            // exportCsvToolStripMenuItem
            exportCsvToolStripMenuItem.Name = "exportCsvToolStripMenuItem";
            exportCsvToolStripMenuItem.Text = "Export CSV...";
            exportCsvToolStripMenuItem.Click += async (_, __) => await ExportCsvAsync();

            // recentFilesToolStripMenuItem
            recentFilesToolStripMenuItem.Name = "recentFilesToolStripMenuItem";
            recentFilesToolStripMenuItem.Text = "Recent Files";

            // openWorkbookToolStripMenuItem
            openWorkbookToolStripMenuItem.Name = "openWorkbookToolStripMenuItem";
            openWorkbookToolStripMenuItem.Text = "Open Workbook...";
            openWorkbookToolStripMenuItem.Click += async (_, __) => await OpenWorkbookAsync();

            // saveWorkbookToolStripMenuItem
            saveWorkbookToolStripMenuItem.Name = "saveWorkbookToolStripMenuItem";
            saveWorkbookToolStripMenuItem.Text = "Save Workbook As...";
            saveWorkbookToolStripMenuItem.Click += async (_, __) => await SaveWorkbookAsync();

            // exitToolStripMenuItem
            exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            exitToolStripMenuItem.Text = "Exit";
            exitToolStripMenuItem.Click += (_, __) => Close();

            // editToolStripMenuItem
            editToolStripMenuItem.Name = "editToolStripMenuItem";
            editToolStripMenuItem.Text = "Edit";
            // Edit menu items (with Fill Down/Right)
            var fillDownToolStripMenuItem = new ToolStripMenuItem("Fill Down") { ShortcutKeys = Keys.Control | Keys.D };
            fillDownToolStripMenuItem.Click += (_, __) => FillDown();
            var fillRightToolStripMenuItem = new ToolStripMenuItem("Fill Right") { ShortcutKeys = Keys.Control | Keys.R };
            fillRightToolStripMenuItem.Click += (_, __) => FillRight();
            editToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] {
                undoToolStripMenuItem, redoToolStripMenuItem,
                new ToolStripSeparator(),
                copyToolStripMenuItem, pasteToolStripMenuItem, cutToolStripMenuItem,
                new ToolStripSeparator(),
                fillDownToolStripMenuItem, fillRightToolStripMenuItem,
                new ToolStripSeparator(),
                findToolStripMenuItem, replaceToolStripMenuItem,
                new ToolStripSeparator(),
                clearContentsToolStripMenuItem, recalcToolStripMenuItem });
            editToolStripMenuItem.DropDownOpening += (_, __) => UpdateAiMenuItemsState();

            // undoToolStripMenuItem
            undoToolStripMenuItem.Name = "undoToolStripMenuItem";
            undoToolStripMenuItem.Text = "Undo";
            undoToolStripMenuItem.ShortcutKeys = Keys.Control | Keys.Z;
            undoToolStripMenuItem.Click += (_, __) => Undo();

            // redoToolStripMenuItem
            redoToolStripMenuItem.Name = "redoToolStripMenuItem";
            redoToolStripMenuItem.Text = "Redo";
            redoToolStripMenuItem.ShortcutKeys = Keys.Control | Keys.Y;
            redoToolStripMenuItem.Click += (_, __) => Redo();

            // copyToolStripMenuItem
            copyToolStripMenuItem.Name = "copyToolStripMenuItem";
            copyToolStripMenuItem.Text = "Copy";
            copyToolStripMenuItem.ShortcutKeys = Keys.Control | Keys.C;
            copyToolStripMenuItem.Click += (_, __) => CopyCell();

            // pasteToolStripMenuItem
            pasteToolStripMenuItem.Name = "pasteToolStripMenuItem";
            pasteToolStripMenuItem.Text = "Paste";
            pasteToolStripMenuItem.ShortcutKeys = Keys.Control | Keys.V;
            pasteToolStripMenuItem.Click += (_, __) => PasteCell();

            // cutToolStripMenuItem
            cutToolStripMenuItem.Name = "cutToolStripMenuItem";
            cutToolStripMenuItem.Text = "Cut";
            cutToolStripMenuItem.ShortcutKeys = Keys.Control | Keys.X;
            cutToolStripMenuItem.Click += (_, __) => CutCell();

            // findToolStripMenuItem
            findToolStripMenuItem.Name = "findToolStripMenuItem";
            findToolStripMenuItem.Text = "Find...";
            findToolStripMenuItem.ShortcutKeys = Keys.Control | Keys.F;
            findToolStripMenuItem.Click += (_, __) => OpenFindDialog(false);

            // replaceToolStripMenuItem
            replaceToolStripMenuItem.Name = "replaceToolStripMenuItem";
            replaceToolStripMenuItem.Text = "Replace...";
            replaceToolStripMenuItem.ShortcutKeys = Keys.Control | Keys.H;
            replaceToolStripMenuItem.Click += (_, __) => OpenFindDialog(true);

            // clearContentsToolStripMenuItem
            clearContentsToolStripMenuItem.Name = "clearContentsToolStripMenuItem";
            clearContentsToolStripMenuItem.Text = "Clear Contents";
            clearContentsToolStripMenuItem.ShortcutKeys = Keys.Delete;
            clearContentsToolStripMenuItem.Click += (_, __) => ClearSelectedCells();

            // recalcToolStripMenuItem
            recalcToolStripMenuItem.Name = "recalcToolStripMenuItem";
            recalcToolStripMenuItem.Text = "Recalculate";
            recalcToolStripMenuItem.ShortcutKeys = Keys.F9;
            recalcToolStripMenuItem.Click += (_, __) => RecalculateAll();

            // formatToolStripMenuItem
            formatToolStripMenuItem.Name = "formatToolStripMenuItem";
            formatToolStripMenuItem.Text = "Format";
            var boldItem = new ToolStripMenuItem("Bold") { CheckOnClick = true };
            boldItem.Click += (_, __) => ToggleBoldFormat();
            var textColorItem = new ToolStripMenuItem("Text Color...");
            textColorItem.Click += (_, __) => ChooseTextColor();
            var fillColorItem = new ToolStripMenuItem("Fill Color...");
            fillColorItem.Click += (_, __) => ChooseFillColor();
            var alignMenu = new ToolStripMenuItem("Align");
            var alignLeft = new ToolStripMenuItem("Left"); alignLeft.Click += (_, __) => SetAlignment("Left");
            var alignCenter = new ToolStripMenuItem("Center"); alignCenter.Click += (_, __) => SetAlignment("Center");
            var alignRight = new ToolStripMenuItem("Right"); alignRight.Click += (_, __) => SetAlignment("Right");
            alignMenu.DropDownItems.AddRange(new ToolStripItem[] { alignLeft, alignCenter, alignRight });
            var numFmtMenu = new ToolStripMenuItem("Number Format");
            var nfGeneral = new ToolStripMenuItem("General"); nfGeneral.Click += (_, __) => SetNumberFormat("General");
            var nf0 = new ToolStripMenuItem("0"); nf0.Click += (_, __) => SetNumberFormat("0");
            var nf2 = new ToolStripMenuItem("0.00"); nf2.Click += (_, __) => SetNumberFormat("0.00");
            var nfT0 = new ToolStripMenuItem("#,##0"); nfT0.Click += (_, __) => SetNumberFormat("#,##0");
            var nfT2 = new ToolStripMenuItem("#,##0.00"); nfT2.Click += (_, __) => SetNumberFormat("#,##0.00");
            var nfP0 = new ToolStripMenuItem("0%"); nfP0.Click += (_, __) => SetNumberFormat("0%");
            var nfP2 = new ToolStripMenuItem("0.00%"); nfP2.Click += (_, __) => SetNumberFormat("0.00%");
            var nfC0 = new ToolStripMenuItem("$#,##0"); nfC0.Click += (_, __) => SetNumberFormat("$#,##0");
            var nfC2 = new ToolStripMenuItem("$#,##0.00"); nfC2.Click += (_, __) => SetNumberFormat("$#,##0.00");
            numFmtMenu.DropDownItems.AddRange(new ToolStripItem[] { nfGeneral, new ToolStripSeparator(), nf0, nf2, new ToolStripSeparator(), nfT0, nfT2, new ToolStripSeparator(), nfP0, nfP2, new ToolStripSeparator(), nfC0, nfC2 });
            formatToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { boldItem, textColorItem, fillColorItem, new ToolStripSeparator(), alignMenu, numFmtMenu });

            // aiToolStripMenuItem
            aiToolStripMenuItem = new ToolStripMenuItem();
            aiToolStripMenuItem.Name = "aiToolStripMenuItem";
            aiToolStripMenuItem.Text = "AI";
            aiGenerateFillToolStripMenuItem = new ToolStripMenuItem("Generate Fill...");
            aiGenerateFillToolStripMenuItem.ShortcutKeys = Keys.Control | Keys.I;
            aiGenerateFillToolStripMenuItem.Click += (_, __) => OpenGenerateFill();
            aiEnableInlineToolStripMenuItem = new ToolStripMenuItem("Enable Inline Suggestions") { CheckOnClick = true, Checked = true };
            aiEnableInlineToolStripMenuItem.CheckedChanged += (_, __) => ToggleInlineSuggestions(aiEnableInlineToolStripMenuItem.Checked);
            aiAcceptInlineToolStripMenuItem = new ToolStripMenuItem("Apply AI Changes") { ShortcutKeys = Keys.Control | Keys.Shift | Keys.I };
            aiAcceptInlineToolStripMenuItem.Click += (_, __) => AcceptInlineSuggestion();
            aiToggleChatPaneToolStripMenuItem = new ToolStripMenuItem("Toggle Docked Chat Pane") { ShortcutKeys = Keys.Control | Keys.Shift | Keys.C };
            aiToggleChatPaneToolStripMenuItem.Click += (_, __) => ToggleChatPane();
            aiOpenChatToolStripMenuItem = new ToolStripMenuItem("Open Chat Assistant (window)...");
            aiOpenChatToolStripMenuItem.Click += (_, __) => OpenChatAssistant();
            aiSettingsToolStripMenuItem = new ToolStripMenuItem("Settings...");
            aiSettingsToolStripMenuItem.Click += (_, __) => OpenAiSettings();
            // AI menu (add Schema Fill)
            var aiSchemaFillToolStripMenuItem = new ToolStripMenuItem("Fill Selected From Schema...");
            aiSchemaFillToolStripMenuItem.Click += async (_, __) => await FillSelectedFromSchemaAsync();
            var aiSmartSchemaFillToolStripMenuItem = new ToolStripMenuItem("Smart Schema Fill") { ShortcutKeys = Keys.Control | Keys.Shift | Keys.F };
            aiSmartSchemaFillToolStripMenuItem.Click += async (_, __) => await SmartSchemaFillAsync();
            var batchSchemaFillToolStripMenuItem = new ToolStripMenuItem("Batch Schema Fill...");
            batchSchemaFillToolStripMenuItem.Click += async (_, __) => await BatchFillSelectedFromSchemaAsync();
            var aiViewActionLogToolStripMenuItem = new ToolStripMenuItem("View Action Log...");
            aiViewActionLogToolStripMenuItem.Click += (_, __) => ShowAIActionLog();
            var aiExplainCellToolStripMenuItem = new ToolStripMenuItem("Explain Cell...");
            aiExplainCellToolStripMenuItem.Click += async (_, __) => await ExplainCellAsync();
            aiToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { aiGenerateFillToolStripMenuItem, aiSchemaFillToolStripMenuItem, aiSmartSchemaFillToolStripMenuItem, batchSchemaFillToolStripMenuItem, aiToggleChatPaneToolStripMenuItem, aiOpenChatToolStripMenuItem, new ToolStripSeparator(), aiEnableInlineToolStripMenuItem, aiAcceptInlineToolStripMenuItem, new ToolStripSeparator(), aiExplainCellToolStripMenuItem, aiViewActionLogToolStripMenuItem, new ToolStripSeparator(), aiSettingsToolStripMenuItem });
            menuStrip1.Items.Add(aiToolStripMenuItem);

            // testToolStripMenuItem
            testToolStripMenuItem.Name = "testToolStripMenuItem";
            testToolStripMenuItem.Text = "Test";
            testRunnerToolStripMenuItem.Name = "testRunnerToolStripMenuItem";
            testRunnerToolStripMenuItem.Text = "Test Runner...";
            testRunnerToolStripMenuItem.Click += (_, __) => OpenTestRunner();
            testToolStripMenuItem.DropDownItems.Add(testRunnerToolStripMenuItem);
            menuStrip1.Items.Add(testToolStripMenuItem);

            // helpToolStripMenuItem
            helpToolStripMenuItem = new ToolStripMenuItem();
            helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            helpToolStripMenuItem.Text = "Help";
            docsViewerToolStripMenuItem = new ToolStripMenuItem("View Docs...");
            docsViewerToolStripMenuItem.Click += (_, __) => OpenDocsViewer();
            exportDocsJsonToolStripMenuItem = new ToolStripMenuItem("Export Docs JSON");
            exportDocsJsonToolStripMenuItem.Click += (_, __) => ExportDocsJson();
            helpToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { docsViewerToolStripMenuItem, exportDocsJsonToolStripMenuItem });
            menuStrip1.Items.Add(helpToolStripMenuItem);

            // viewToolStripMenuItem (View menu with Freeze options)
            var viewToolStripMenuItem = new ToolStripMenuItem();
            viewToolStripMenuItem.Name = "viewToolStripMenuItem";
            viewToolStripMenuItem.Text = "View";
            var freezeTopRowToolStripMenuItem = new ToolStripMenuItem("Freeze Top Row") { CheckOnClick = true };
            freezeTopRowToolStripMenuItem.Click += (_, __) => ToggleFreezeTopRow();
            var freezeFirstColToolStripMenuItem = new ToolStripMenuItem("Freeze First Column") { CheckOnClick = true };
            freezeFirstColToolStripMenuItem.Click += (_, __) => ToggleFreezeFirstColumn();
            viewToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { freezeTopRowToolStripMenuItem, freezeFirstColToolStripMenuItem });
            menuStrip1.Items.Insert(menuStrip1.Items.IndexOf(helpToolStripMenuItem), viewToolStripMenuItem);

            // sheetsToolStripMenuItem
            sheetsToolStripMenuItem.Name = "sheetsToolStripMenuItem";
            sheetsToolStripMenuItem.Text = "Sheets";
            sheetsToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { addSheetToolStripMenuItem, renameSheetToolStripMenuItem, removeSheetToolStripMenuItem });

            addSheetToolStripMenuItem.Name = "addSheetToolStripMenuItem";
            addSheetToolStripMenuItem.Text = "Add Sheet";
            addSheetToolStripMenuItem.Click += (_, __) => AddSheet();

            renameSheetToolStripMenuItem.Name = "renameSheetToolStripMenuItem";
            renameSheetToolStripMenuItem.Text = "Rename Sheet";
            renameSheetToolStripMenuItem.Click += (_, __) => RenameSheet();

            removeSheetToolStripMenuItem.Name = "removeSheetToolStripMenuItem";
            removeSheetToolStripMenuItem.Text = "Remove Sheet";
            removeSheetToolStripMenuItem.Click += (_, __) => RemoveSheet();

            // tabs
            tabs.Dock = DockStyle.Top;
            tabs.Height = 26;
            tabs.SelectedIndexChanged += Tabs_SelectedIndexChanged;

            // Formula bar
            _formulaBarPanel = new Panel();
            _formulaBarPanel.Dock = DockStyle.Top;
            _formulaBarPanel.Height = 26;
            _formulaBarPanel.BackColor = Color.FromArgb(240, 240, 240);
            _formulaBarPanel.Padding = new Padding(0);

            _cellNameBox = new TextBox();
            _cellNameBox.Dock = DockStyle.Left;
            _cellNameBox.Width = 60;
            _cellNameBox.Font = new Font("Segoe UI", 9F);
            _cellNameBox.TextAlign = HorizontalAlignment.Center;
            _cellNameBox.BorderStyle = BorderStyle.FixedSingle;

            var formulaSep = new Label();
            formulaSep.Dock = DockStyle.Left;
            formulaSep.Width = 1;
            formulaSep.BackColor = Color.FromArgb(200, 200, 200);

            _formulaBar = new TextBox();
            _formulaBar.Dock = DockStyle.Fill;
            _formulaBar.Font = new Font("Segoe UI", 9F);
            _formulaBar.BackColor = Color.White;
            _formulaBar.BorderStyle = BorderStyle.FixedSingle;

            _formulaBarPanel.Controls.Add(_formulaBar);     // Fill — added first
            _formulaBarPanel.Controls.Add(formulaSep);      // Left
            _formulaBarPanel.Controls.Add(_cellNameBox);    // Left

            // grid
            grid.Dock = DockStyle.Fill;
            grid.Location = new System.Drawing.Point(0, 24);
            grid.Name = "grid";
            grid.AllowUserToAddRows = false;
            grid.AllowUserToDeleteRows = false;
            grid.AllowUserToResizeRows = true;
            grid.MultiSelect = true;
            grid.RowHeadersWidth = 60;
            grid.SelectionMode = DataGridViewSelectionMode.CellSelect;
            grid.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
            grid.TabIndex = 1;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            grid.AllowUserToResizeColumns = true;

            // Modern flat styling
            grid.BorderStyle = BorderStyle.None;
            grid.CellBorderStyle = DataGridViewCellBorderStyle.Single;
            grid.GridColor = Color.FromArgb(228, 228, 228);
            grid.BackgroundColor = Color.White;

            // Default cell style
            grid.DefaultCellStyle.Font = new Font("Segoe UI", 9F);
            grid.DefaultCellStyle.BackColor = Color.White;
            grid.DefaultCellStyle.ForeColor = Color.FromArgb(30, 30, 30);
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(200, 220, 240);
            grid.DefaultCellStyle.SelectionForeColor = Color.FromArgb(30, 30, 30);
            grid.DefaultCellStyle.Padding = new Padding(2, 0, 2, 0);

            // Flat column headers
            grid.EnableHeadersVisualStyles = false;
            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(240, 240, 240);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(60, 60, 60);
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F);
            grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.FromArgb(240, 240, 240);
            grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = Color.FromArgb(60, 60, 60);
            grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            grid.ColumnHeadersHeight = 28;

            // Flat row headers
            grid.RowHeadersDefaultCellStyle.BackColor = Color.FromArgb(240, 240, 240);
            grid.RowHeadersDefaultCellStyle.ForeColor = Color.FromArgb(60, 60, 60);
            grid.RowHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F);
            grid.RowHeadersDefaultCellStyle.SelectionBackColor = Color.FromArgb(240, 240, 240);
            grid.RowHeadersDefaultCellStyle.SelectionForeColor = Color.FromArgb(60, 60, 60);
            grid.RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;

            grid.CellBeginEdit += Grid_CellBeginEdit;
            grid.CellEndEdit += Grid_CellEndEdit;
            grid.SelectionChanged += Grid_SelectionChanged;
            grid.KeyDown += Grid_KeyDown;
            grid.DataError += (s, e) => { e.ThrowException = false; };
            grid.Paint += (s, e) => Grid_Paint(e);
            grid.MouseDown += Grid_MouseDown;
            grid.MouseMove += Grid_MouseMove;
            grid.MouseUp += Grid_MouseUp;
            grid.MouseDoubleClick += Grid_MouseDoubleClick;

            // statusStrip1
            statusStrip1.Dock = DockStyle.Bottom;
            statusStrip1.Name = "statusStrip1";
            statusStrip1.SizingGrip = false;
            statusCell = new ToolStripStatusLabel { Name = "statusCell", Text = "Cell: -" };
            statusRaw = new ToolStripStatusLabel { Name = "statusRaw", Text = "Raw: " };
            statusValue = new ToolStripStatusLabel { Name = "statusValue", Text = "Value: ", Spring = true, TextAlign = System.Drawing.ContentAlignment.MiddleLeft };
            statusStrip1.Items.AddRange(new ToolStripItem[] { statusCell, statusRaw, statusValue });

            // MainForm
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(1000, 700);
            Controls.Add(grid);              // Fill
            Controls.Add(tabs);              // Top
            Controls.Add(_formulaBarPanel);  // Top — between tabs and menu
            Controls.Add(statusStrip1);      // Bottom
            Controls.Add(menuStrip1);        // Top — topmost
            MainMenuStrip = menuStrip1;
            Name = "MainForm";
            Text = "Spreadsheet";
            Load += (_, __) => InitializeSheet();

            ResumeLayout(false);
            PerformLayout();
        }
    }
}
