using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using SpreadsheetApp.Core;

namespace SpreadsheetApp.UI
{
    public sealed class DocsViewerForm : Form
    {
        private readonly ListBox _lstFiles = new() { Dock = DockStyle.Fill };
        private readonly ListBox _lstSections = new() { Dock = DockStyle.Fill };
        private readonly WebBrowser _browser = new() { Dock = DockStyle.Fill, AllowWebBrowserDrop = false, ScriptErrorsSuppressed = true };
        private readonly Button _btnExport = new() { Text = "Export JSON", Width = 110, Height = 28 };
        private readonly Label _lblStatus = new() { AutoSize = true };

        private DocsIndex _index = new();
        private readonly string _root;

        public DocsViewerForm()
        {
            Text = "Docs Viewer";
            StartPosition = FormStartPosition.CenterParent;
            Size = new Size(960, 640);
            MinimumSize = new Size(720, 480);
            FormBorderStyle = FormBorderStyle.Sizable;

            _root = DocsIndexer.TryFindRepoRoot() ?? Environment.CurrentDirectory;
            BuildLayout();
            LoadIndex();

            _lstFiles.SelectedIndexChanged += (_, __) => RefreshSections();
            _lstSections.SelectedIndexChanged += (_, __) => RefreshContent();
            _btnExport.Click += (_, __) => ExportJson();
        }

        private void BuildLayout()
        {
            var top = new Panel { Dock = DockStyle.Top, Height = 36, Padding = new Padding(8, 6, 8, 6) };
            _btnExport.Location = new Point(8, 4);
            _lblStatus.Location = new Point(130, 8);
            top.Controls.Add(_btnExport);
            top.Controls.Add(_lblStatus);

            var leftSplit = new SplitContainer { Dock = DockStyle.Left, Width = 320, Orientation = Orientation.Horizontal, SplitterDistance = 240 };
            leftSplit.Panel1.Padding = new Padding(6);
            leftSplit.Panel2.Padding = new Padding(6);
            leftSplit.Panel1.Controls.Add(_lstFiles);
            leftSplit.Panel2.Controls.Add(_lstSections);

            var contentPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(6) };
            contentPanel.Controls.Add(_browser);

            Controls.Add(contentPanel);
            Controls.Add(leftSplit);
            Controls.Add(top);
        }

        private void LoadIndex()
        {
            try
            {
                _index = DocsIndexer.Build(_root);
                _lstFiles.Items.Clear();
                foreach (var f in _index.Files)
                {
                    _lstFiles.Items.Add(f.Title + "  (" + f.Name + ")");
                }
                _lblStatus.Text = _index.Files.Count == 0 ? "No markdown files found." : $"{_index.Files.Count} docs found";
                if (_lstFiles.Items.Count > 0) _lstFiles.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                _lblStatus.Text = "Failed to load docs: " + ex.Message;
            }
        }

        private void RefreshSections()
        {
            _lstSections.Items.Clear();
            _browser.DocumentText = "";
            int idx = _lstFiles.SelectedIndex;
            if (idx < 0 || idx >= _index.Files.Count) return;
            var file = _index.Files[idx];
            foreach (var s in file.Sections)
            {
                var indent = new string(' ', Math.Max(0, (s.Level - 1) * 2));
                _lstSections.Items.Add(indent + new string('#', s.Level) + " " + s.Title);
            }
            if (_lstSections.Items.Count > 0) _lstSections.SelectedIndex = 0; else RenderMarkdown(file.Title ?? file.Name, file.Content);
        }

        private void RefreshContent()
        {
            _browser.DocumentText = "";
            int fIdx = _lstFiles.SelectedIndex;
            if (fIdx < 0 || fIdx >= _index.Files.Count) return;
            var file = _index.Files[fIdx];
            int sIdx = _lstSections.SelectedIndex;
            string md;
            string title;
            if (sIdx < 0 || sIdx >= file.Sections.Count)
            {
                md = file.Content;
                title = file.Title ?? file.Name;
            }
            else
            {
                var s = file.Sections[sIdx];
                title = s.Title;
                md = new string('#', Math.Max(1, Math.Min(6, s.Level))) + " " + s.Title + "\r\n\r\n" + s.Content;
            }
            RenderMarkdown(title, md);
        }

        private void ExportJson()
        {
            try
            {
                var index = DocsIndexer.Build(_root);
                string path = DocsIndexer.WriteJson(index, null);
                MessageBox.Show(this, $"Exported to:\r\n{path}", "Docs Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Failed to export: " + ex.Message, "Docs Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RenderMarkdown(string title, string md)
        {
            try
            {
                string body = Markdown.ToHtml(md);
                string html = Markdown.WrapHtml(title, body);
                _browser.DocumentText = html;
            }
            catch
            {
                _browser.DocumentText = "<html><body><pre>Failed to render markdown.</pre></body></html>";
            }
        }
    }
}
