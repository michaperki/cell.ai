using System;
using System.Drawing;
using System.Windows.Forms;

namespace SpreadsheetApp.UI
{
    public class FindReplaceForm : Form
    {
        private TextBox _findBox = null!;
        private TextBox _replaceBox = null!;
        private CheckBox _matchCase = null!;
        private Button _btnFindNext = null!;
        private Button _btnReplace = null!;
        private Button _btnReplaceAll = null!;
        private Button _btnClose = null!;

        public event EventHandler? FindNextClicked;
        public event EventHandler? ReplaceClicked;
        public event EventHandler? ReplaceAllClicked;

        public string FindText
        {
            get => _findBox.Text;
            set => _findBox.Text = value ?? string.Empty;
        }

        public string ReplaceText
        {
            get => _replaceBox.Text;
            set => _replaceBox.Text = value ?? string.Empty;
        }

        public bool MatchCase
        {
            get => _matchCase.Checked;
            set => _matchCase.Checked = value;
        }

        public FindReplaceForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            Text = "Find / Replace";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ClientSize = new Size(420, 160);

            var lblFind = new Label { Text = "Find:", AutoSize = true, Location = new Point(12, 16) };
            _findBox = new TextBox { Location = new Point(80, 12), Width = 320 };
            var lblReplace = new Label { Text = "Replace:", AutoSize = true, Location = new Point(12, 48) };
            _replaceBox = new TextBox { Location = new Point(80, 44), Width = 320 };
            _matchCase = new CheckBox { Text = "Match case", AutoSize = true, Location = new Point(80, 76) };

            _btnFindNext = new Button { Text = "Find Next", Location = new Point(12, 110), Size = new Size(90, 28) };
            _btnReplace = new Button { Text = "Replace", Location = new Point(108, 110), Size = new Size(90, 28) };
            _btnReplaceAll = new Button { Text = "Replace All", Location = new Point(204, 110), Size = new Size(90, 28) };
            _btnClose = new Button { Text = "Close", Location = new Point(310, 110), Size = new Size(90, 28) };

            _btnFindNext.Click += (s, e) => FindNextClicked?.Invoke(this, EventArgs.Empty);
            _btnReplace.Click += (s, e) => ReplaceClicked?.Invoke(this, EventArgs.Empty);
            _btnReplaceAll.Click += (s, e) => ReplaceAllClicked?.Invoke(this, EventArgs.Empty);
            _btnClose.Click += (s, e) => Close();

            Controls.AddRange(new Control[] { lblFind, _findBox, lblReplace, _replaceBox, _matchCase, _btnFindNext, _btnReplace, _btnReplaceAll, _btnClose });

            AcceptButton = _btnFindNext;
            CancelButton = _btnClose;
        }
    }
}

