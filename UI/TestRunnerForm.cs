using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace SpreadsheetApp.UI
{
    public sealed class TestRunnerForm : Form
    {
        private readonly Action<string> _loadWorkbook;
        private readonly Func<string, string?, bool, System.Threading.CancellationToken, System.Threading.Tasks.Task<SpreadsheetApp.Core.AI.AIPlan>> _runChatStepAsync;
        private readonly Func<string, string?, bool, System.Threading.CancellationToken, System.Threading.Tasks.Task<(SpreadsheetApp.Core.AI.AIPlan plan, string[] transcript)>> _runAgentStepAsync;
        private readonly Action<string> _saveWorkbook;
        private readonly Action<string> _activateSheet;
        private readonly Action _clearChatHistory;
        private readonly Func<System.Collections.Generic.Dictionary<string, string>> _captureSheetMap;
        private readonly ListBox _lstTests = new();
        private readonly TextBox _txtSpec = new() { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, BackColor = SystemColors.Info, Font = new Font("Segoe UI", 9.5f) };
        private readonly TextBox _txtLog = new() { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, BackColor = SystemColors.Window, Font = new Font("Consolas", 9f) };
        private readonly Button _btnPrev = new() { Text = "<< Prev", Width = 80, Height = 30 };
        private readonly Button _btnNext = new() { Text = "Next >>", Width = 80, Height = 30 };
        private readonly Button _btnLoad = new() { Text = "Load Test", Width = 100, Height = 30 };
        private readonly Button _btnRun = new() { Text = "Run Steps", Width = 100, Height = 30 };
        private readonly CheckBox _chkSnapshots = new() { Text = "Save snapshots", AutoSize = true, Checked = true };
        private readonly CheckBox _chkDumpPlan = new() { Text = "Dump plan JSON", AutoSize = true, Checked = true };
        private readonly CheckBox _chkDumpPrompt = new() { Text = "Dump user prompt", AutoSize = true, Checked = true };
        private readonly CheckBox _chkDumpSystem = new() { Text = "Dump system prompt", AutoSize = true, Checked = true };
        private readonly Label _lblStatus = new() { AutoSize = true, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
        private readonly List<string> _testFiles = new();
        private readonly List<string> _visibleFiles = new();
        private readonly ComboBox _cmbFilter = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 140 };
        private TestSpecs? _specs;

        public TestRunnerForm(Action<string> loadWorkbook,
                              Func<string, string?, bool, System.Threading.CancellationToken, System.Threading.Tasks.Task<SpreadsheetApp.Core.AI.AIPlan>> runChatStepAsync,
                              Func<string, string?, bool, System.Threading.CancellationToken, System.Threading.Tasks.Task<(SpreadsheetApp.Core.AI.AIPlan plan, string[] transcript)>> runAgentStepAsync,
                              Action<string> saveWorkbook,
                              Action<string> activateSheet,
                              Action clearChatHistory,
                              Func<System.Collections.Generic.Dictionary<string, string>> captureSheetMap)
        {
            _loadWorkbook = loadWorkbook;
            _runChatStepAsync = runChatStepAsync;
            _runAgentStepAsync = runAgentStepAsync;
            _saveWorkbook = saveWorkbook;
            _activateSheet = activateSheet;
            _clearChatHistory = clearChatHistory;
            _captureSheetMap = captureSheetMap;
            Text = "Test Runner";
            StartPosition = FormStartPosition.CenterParent;
            Size = new Size(760, 600);
            MinimumSize = new Size(640, 480);
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = false;

            DiscoverTests();
            LoadSpecs();
            BuildLayout();

            _lstTests.SelectedIndexChanged += (_, __) => UpdatePreview();
            _btnPrev.Click += (_, __) => Navigate(-1);
            _btnNext.Click += (_, __) => Navigate(1);
            _btnLoad.Click += (_, __) => LoadSelected();
            _btnRun.Click += async (_, __) => await RunSelectedAsync();
            _cmbFilter.SelectedIndexChanged += (_, __) => RefreshVisibleList();

            RefreshVisibleList();
            if (_lstTests.Items.Count > 0) _lstTests.SelectedIndex = 0;
        }

        private void DiscoverTests()
        {
            // Look for tests/ directory relative to the executable, then relative to common dev paths
            string? testsDir = FindTestsDir();
            if (testsDir == null || !Directory.Exists(testsDir))
            {
                _lblStatus.Text = "No tests/ directory found.";
                return;
            }

            var files = Directory.GetFiles(testsDir, "test_*.workbook.json")
                .OrderBy(f => Path.GetFileName(f))
                .ToList();

            foreach (var f in files) _testFiles.Add(f);
            _lblStatus.Text = $"{_testFiles.Count} tests found";
        }

        private static string? FindTestsDir()
        {
            // Try relative to the app base directory
            string appBase = AppDomain.CurrentDomain.BaseDirectory;
            // Walk up from bin/Debug/net8.0-windows to project root
            var dir = new DirectoryInfo(appBase);
            for (int i = 0; i < 6 && dir != null; i++)
            {
                string candidate = Path.Combine(dir.FullName, "tests");
                if (Directory.Exists(candidate)) return candidate;
                dir = dir.Parent;
            }
            // Fallback: current working directory
            string cwd = Path.Combine(Directory.GetCurrentDirectory(), "tests");
            if (Directory.Exists(cwd)) return cwd;
            return null;
        }

        private void BuildLayout()
        {
            var topPanel = new Panel { Dock = DockStyle.Top, Height = 36, Padding = new Padding(6, 4, 6, 4) };
            _lblStatus.Location = new Point(8, 8);
            topPanel.Controls.Add(_lblStatus);
            _cmbFilter.Items.AddRange(new object[] { "All tests", "AI tests", "Non-AI tests" });
            _cmbFilter.SelectedIndex = 0;
            _cmbFilter.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _cmbFilter.Location = new Point(Width - 200, 6);
            topPanel.Controls.Add(_cmbFilter);

            var listPanel = new Panel { Dock = DockStyle.Top, Height = 180, Padding = new Padding(6, 0, 6, 4) };
            _lstTests.Dock = DockStyle.Fill;
            _lstTests.Font = new Font("Consolas", 9.5f);
            _lstTests.IntegralHeight = false;
            listPanel.Controls.Add(_lstTests);

            var instructionsLabel = new Label { Text = "Test Steps (from TEST_SPECS.json):", Dock = DockStyle.Top, Height = 20, Padding = new Padding(6, 4, 0, 0), Font = new Font("Segoe UI", 8.5f, FontStyle.Italic) };

            var split = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Horizontal, SplitterDistance = 160 };
            var instrPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(6, 0, 6, 4) };
            _txtSpec.Dock = DockStyle.Fill;
            instrPanel.Controls.Add(_txtSpec);

            var logPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(6, 0, 6, 6) };
            var logLabel = new Label { Text = "Log:", Dock = DockStyle.Top, Height = 18, Padding = new Padding(0, 0, 0, 2) };
            _txtLog.Dock = DockStyle.Fill;
            logPanel.Controls.Add(_txtLog);
            logPanel.Controls.Add(logLabel);

            split.Panel1.Controls.Add(instrPanel);
            split.Panel2.Controls.Add(logPanel);

            var btnPanel = new Panel { Dock = DockStyle.Bottom, Height = 44, Padding = new Padding(6, 4, 6, 6) };
            _btnPrev.Location = new Point(6, 4);
            _btnNext.Location = new Point(92, 4);
            _btnLoad.Location = new Point(200, 4);
            _btnRun.Location = new Point(320, 4);
            _chkSnapshots.Location = new Point(440, 10);
            _chkDumpPlan.Location = new Point(560, 10);
            _chkDumpPrompt.Location = new Point(690, 10);
            _chkDumpSystem.Location = new Point(840, 10);
            btnPanel.Controls.AddRange(new Control[] { _btnPrev, _btnNext, _btnLoad, _btnRun, _chkSnapshots, _chkDumpPlan, _chkDumpPrompt, _chkDumpSystem });

            Controls.Add(split);
            Controls.Add(instructionsLabel);
            Controls.Add(listPanel);
            Controls.Add(topPanel);
            Controls.Add(btnPanel);
        }

        private void Navigate(int delta)
        {
            if (_lstTests.Items.Count == 0) return;
            int idx = _lstTests.SelectedIndex + delta;
            if (idx < 0) idx = 0;
            if (idx >= _lstTests.Items.Count) idx = _lstTests.Items.Count - 1;
            _lstTests.SelectedIndex = idx;
        }

        private void UpdatePreview()
        {
            int idx = _lstTests.SelectedIndex;
            if (idx < 0 || idx >= _visibleFiles.Count)
            {
                _txtSpec.Text = string.Empty; _txtLog.Clear();
                return;
            }

            _lblStatus.Text = $"Test {idx + 1} of {_visibleFiles.Count}";
            _btnPrev.Enabled = idx > 0;
            _btnNext.Enabled = idx < _visibleFiles.Count - 1;
            _txtLog.Clear();
            _txtSpec.Text = DescribeSpecForPath(_visibleFiles[idx]);
        }

        private void LoadSelected()
        {
            int idx = _lstTests.SelectedIndex;
            if (idx < 0 || idx >= _visibleFiles.Count) return;
            try
            {
                _loadWorkbook(_visibleFiles[idx]);
                _lblStatus.Text = $"Test {idx + 1} of {_visibleFiles.Count} -- LOADED";
                _clearChatHistory();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Failed to load test: {ex.Message}", "Test Runner", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Log(string text)
        {
            if (IsDisposed || _txtLog.IsDisposed) return;
            _txtLog.AppendText(text);
        }

        private async System.Threading.Tasks.Task RunSelectedAsync()
        {
            int idx = _lstTests.SelectedIndex;
            if (idx < 0 || idx >= _visibleFiles.Count) return;
            string path = _visibleFiles[idx];
            var spec = GetSpecForPath(path);
            if (spec == null || spec.Steps == null || spec.Steps.Count == 0)
            {
                MessageBox.Show(this, "No steps defined for this test in TEST_SPECS.json.", "Test Runner", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            Log($"> Running {Path.GetFileName(path)}\r\n");
            _clearChatHistory();
            int stepNo = 1;
            foreach (var step in spec.Steps)
            {
                if (IsDisposed) return;
                try
                {
                    if (!string.IsNullOrWhiteSpace(step.Sheet))
                    {
                        _activateSheet(step.Sheet!);
                        Log($"  - Activated sheet: {step.Sheet}\r\n");
                    }
                    if (step.Action == "ai_chat")
                    {
                        Log($"  - Step {stepNo}: prompt=\"{step.Prompt}\" loc={(step.Location ?? "(none)")} apply={step.Apply}\r\n");
                        var before = _captureSheetMap();
                        var plan = await _runChatStepAsync(step.Prompt ?? string.Empty, step.Location, step.Apply, System.Threading.CancellationToken.None).ConfigureAwait(true);
                        if (plan.Commands.Count == 0)
                        {
                            Log("    -> No changes suggested.\r\n");
                        }
                        else
                        {
                            foreach (var cmd in plan.Commands) Log($"    -> {cmd.Summarize()}\r\n");
                        }
                        // Diff (only when apply)
                        if (step.Apply)
                        {
                            var after = _captureSheetMap();
                            WriteDiffToLog(before, after);
                            // Structural assertion: all changes must be within the requested selection (if provided)
                            try
                            {
                                if (!string.IsNullOrWhiteSpace(step.Location) && step.Location!.Contains(":", StringComparison.Ordinal))
                                {
                                    var parts = step.Location!.Split(':');
                                    if (parts.Length == 2 && SpreadsheetApp.Core.CellAddress.TryParse(parts[0].Trim(), out int r1, out int c1) && SpreadsheetApp.Core.CellAddress.TryParse(parts[1].Trim(), out int r2, out int c2))
                                    {
                                        int rStart = Math.Min(r1, r2), rEnd = Math.Max(r1, r2);
                                        int cStart = Math.Min(c1, c2), cEnd = Math.Max(c1, c2);
                                        var keys = new System.Collections.Generic.HashSet<string>(before.Keys, System.StringComparer.OrdinalIgnoreCase);
                                        foreach (var k in after.Keys) keys.Add(k);
                                        foreach (var k in keys)
                                        {
                                            before.TryGetValue(k, out var b);
                                            after.TryGetValue(k, out var a);
                                            if (string.Equals(b ?? string.Empty, a ?? string.Empty, StringComparison.Ordinal)) continue;
                                            // parse address
                                            if (SpreadsheetApp.Core.CellAddress.TryParse(k, out int rr, out int cc))
                                            {
                                                if (rr < rStart || rr > rEnd || cc < cStart || cc > cEnd)
                                                {
                                                    Log($"    ASSERTION FAILED: change outside selection: {k}\r\n");
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            catch { }
                        }
                        // Always ensure a snapshot exists for consolidated export; save once here
                        try
                        {
                            Directory.CreateDirectory(Path.Combine("tests", "output"));
                            string baseName = Path.GetFileNameWithoutExtension(path);
                            string snap = Path.Combine("tests", "output", $"{baseName}_step{stepNo}.workbook.json");
                            _saveWorkbook(snap);
                            if (_chkSnapshots.Checked)
                            {
                                Log($"    -> Snapshot saved: {snap}\r\n");
                            }
                        }
                        catch (Exception ex)
                        {
                            Log($"    -> Failed to save snapshot: {ex.Message}\r\n");
                        }

                        if (_chkDumpPrompt.Checked)
                        {
                            try
                            {
                                Directory.CreateDirectory(Path.Combine("tests", "output"));
                                string baseName = Path.GetFileNameWithoutExtension(path);
                                string promptPath = Path.Combine("tests", "output", $"{baseName}_step{stepNo}.user.txt");
                                string contents = plan.RawUser ?? $"(no RawUser from provider)\nstepPrompt: {step.Prompt}\nlocation: {step.Location}\n";
                                File.WriteAllText(promptPath, contents);
                                Log($"    -> User prompt saved: {promptPath}\r\n");
                        }
                        catch (Exception ex)
                        {
                            Log($"    -> Failed to save user prompt: {ex.Message}\r\n");
                        }
                    }
                    if (_chkDumpPlan.Checked)
                    {
                            try
                            {
                                Directory.CreateDirectory(Path.Combine("tests", "output"));
                                string baseName = Path.GetFileNameWithoutExtension(path);
                                string planPath = Path.Combine("tests", "output", $"{baseName}_step{stepNo}.plan.json");
                                string contents = string.IsNullOrWhiteSpace(plan.RawJson) ? SerializePlan(plan) : plan.RawJson!;
                                File.WriteAllText(planPath, contents);
                                Log($"    -> Plan JSON saved: {planPath}\r\n");
                            }
                            catch (Exception ex)
                            {
                                Log($"    -> Failed to save plan JSON: {ex.Message}\r\n");
                            }
                        }

                        // Persist a concise step log for offline review
                        try
                        {
                            Directory.CreateDirectory(Path.Combine("tests", "output"));
                            string baseName = Path.GetFileNameWithoutExtension(path);
                            string logPath = Path.Combine("tests", "output", $"{baseName}_step{stepNo}.log.txt");
                            File.WriteAllText(logPath, _txtLog.Text);
                            Log($"    -> Step log saved: {logPath}\r\n");
                        }
                        catch { }
                        if (_chkDumpSystem.Checked)
                        {
                            try
                            {
                                Directory.CreateDirectory(Path.Combine("tests", "output"));
                                string baseName = Path.GetFileNameWithoutExtension(path);
                                string sysPath = Path.Combine("tests", "output", $"{baseName}_step{stepNo}.system.txt");
                                string contents = plan.RawSystem ?? "(no system prompt)";
                                File.WriteAllText(sysPath, contents);
                                Log($"    -> System prompt saved: {sysPath}\r\n");
                            }
                            catch (Exception ex)
                            {
                                Log($"    -> Failed to save system prompt: {ex.Message}\r\n");
                            }
                        }
                        // Consolidated export (single JSON file containing system, user, plan, and workbook)
                        try
                        {
                            Directory.CreateDirectory(Path.Combine("tests", "output"));
                            string baseName = Path.GetFileNameWithoutExtension(path);
                            string exportPath = Path.Combine("tests", "output", $"{baseName}_step{stepNo}.export.json");
                            string userStr = plan.RawUser ?? $"(no RawUser from provider)\nstepPrompt: {step.Prompt}\nlocation: {step.Location}\n";
                            string sysStr = plan.RawSystem ?? "(no system prompt)";
                            string planStr = string.IsNullOrWhiteSpace(plan.RawJson) ? SerializePlan(plan) : plan.RawJson!;
                            string snapPath = Path.Combine("tests", "output", $"{baseName}_step{stepNo}.workbook.json");
                            string workbookStr = "";
                            try { workbookStr = File.ReadAllText(snapPath); } catch { workbookStr = ""; }
                            var root = new System.Collections.Generic.Dictionary<string, object?>
                            {
                                ["test_file"] = Path.GetFileName(path),
                                ["step"] = stepNo,
                                ["sheet"] = step.Sheet ?? "(current)",
                                ["location"] = step.Location,
                                ["apply"] = step.Apply,
                                ["user"] = userStr,
                                ["system"] = sysStr,
                                ["plan_json"] = planStr,
                                ["workbook_json"] = workbookStr
                            };
                            var opts = new System.Text.Json.JsonSerializerOptions { WriteIndented = true };
                            File.WriteAllText(exportPath, System.Text.Json.JsonSerializer.Serialize(root, opts));
                            Log($"    -> Consolidated export saved: {exportPath}\r\n");
                        }
                        catch (Exception ex)
                        {
                            Log($"    -> Failed to save consolidated export: {ex.Message}\r\n");
                        }
                    }
                    else if (step.Action == "ai_agent")
                    {
                        Log($"  - Step {stepNo} [agent]: prompt=\"{step.Prompt}\" loc={(step.Location ?? "(none)")} apply={step.Apply}\r\n");
                        var before = _captureSheetMap();
                        var res = await _runAgentStepAsync(step.Prompt ?? string.Empty, step.Location, step.Apply, System.Threading.CancellationToken.None).ConfigureAwait(true);
                        var plan = res.plan;
                        var transcript = res.transcript ?? Array.Empty<string>();
                        if (transcript.Length > 0)
                        {
                            Log("    Observations:\r\n");
                            foreach (var line in transcript) Log($"      - {line}\r\n");
                        }
                        if (plan.Commands.Count == 0)
                        {
                            Log("    -> No changes suggested.\r\n");
                        }
                        else
                        {
                            foreach (var cmd in plan.Commands) Log($"    -> {cmd.Summarize()}\r\n");
                        }
                        if (step.Apply)
                        {
                            var after = _captureSheetMap();
                            WriteDiffToLog(before, after);
                            try
                            {
                                if (!string.IsNullOrWhiteSpace(step.Location) && step.Location!.Contains(":", StringComparison.Ordinal))
                                {
                                    var parts = step.Location!.Split(':');
                                    if (parts.Length == 2 && SpreadsheetApp.Core.CellAddress.TryParse(parts[0].Trim(), out int r1, out int c1) && SpreadsheetApp.Core.CellAddress.TryParse(parts[1].Trim(), out int r2, out int c2))
                                    {
                                        int rStart = Math.Min(r1, r2), rEnd = Math.Max(r1, r2);
                                        int cStart = Math.Min(c1, c2), cEnd = Math.Max(c1, c2);
                                        var keys = new System.Collections.Generic.HashSet<string>(before.Keys, System.StringComparer.OrdinalIgnoreCase);
                                        foreach (var k in after.Keys) keys.Add(k);
                                        foreach (var k in keys)
                                        {
                                            before.TryGetValue(k, out var b);
                                            after.TryGetValue(k, out var a);
                                            if (string.Equals(b ?? string.Empty, a ?? string.Empty, StringComparison.Ordinal)) continue;
                                            if (SpreadsheetApp.Core.CellAddress.TryParse(k, out int rr, out int cc))
                                            {
                                                if (rr < rStart || rr > rEnd || cc < cStart || cc > cEnd)
                                                {
                                                    Log($"    ASSERTION FAILED: change outside selection: {k}\r\n");
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            catch { }
                        }
                        try
                        {
                            Directory.CreateDirectory(Path.Combine("tests", "output"));
                            string baseName = Path.GetFileNameWithoutExtension(path);
                            string snap = Path.Combine("tests", "output", $"{baseName}_step{stepNo}.workbook.json");
                            _saveWorkbook(snap);
                            if (_chkSnapshots.Checked)
                            {
                                string logPath = Path.Combine("tests", "output", $"{baseName}_step{stepNo}.log.txt");
                                File.WriteAllText(logPath, _txtLog.Text);
                                Log($"    -> Step log saved: {logPath}\r\n");
                            }
                        }
                        catch { }
                        if (_chkDumpPlan.Checked)
                        {
                            try
                            {
                                Directory.CreateDirectory(Path.Combine("tests", "output"));
                                string baseName = Path.GetFileNameWithoutExtension(path);
                                string planPath = Path.Combine("tests", "output", $"{baseName}_step{stepNo}.plan.json");
                                string json = SerializePlan(plan);
                                File.WriteAllText(planPath, json);
                                Log($"    -> Plan JSON saved: {planPath}\r\n");
                            }
                            catch { }
                        }
                        if (_chkDumpPrompt.Checked)
                        {
                            try
                            {
                                Directory.CreateDirectory(Path.Combine("tests", "output"));
                                string baseName = Path.GetFileNameWithoutExtension(path);
                                string logPath = Path.Combine("tests", "output", $"{baseName}_step{stepNo}.agent.txt");
                                var lines = transcript.Length > 0 ? string.Join("\n", transcript) : "(no transcript)";
                                File.WriteAllText(logPath, lines);
                                Log($"    -> Agent transcript saved: {logPath}\r\n");
                            }
                            catch { }
                        }
                        try
                        {
                            Directory.CreateDirectory(Path.Combine("tests", "output"));
                            string baseName = Path.GetFileNameWithoutExtension(path);
                            string exportPath = Path.Combine("tests", "output", $"{baseName}_step{stepNo}.export.json");
                            string userStr = plan.RawUser ?? $"(no RawUser from provider)\nstepPrompt: {step.Prompt}\nlocation: {step.Location}\n";
                            string sysStr = plan.RawSystem ?? "(no system prompt)";
                            string planStr = string.IsNullOrWhiteSpace(plan.RawJson) ? SerializePlan(plan) : plan.RawJson!;
                            string snapPath = Path.Combine("tests", "output", $"{baseName}_step{stepNo}.workbook.json");
                            string workbookStr = "";
                            try { workbookStr = File.ReadAllText(snapPath); } catch { workbookStr = ""; }
                            var root = new System.Collections.Generic.Dictionary<string, object?>
                            {
                                ["test_file"] = Path.GetFileName(path),
                                ["step"] = stepNo,
                                ["sheet"] = step.Sheet ?? "(current)",
                                ["location"] = step.Location,
                                ["apply"] = step.Apply,
                                ["user"] = userStr,
                                ["system"] = sysStr,
                                ["plan_json"] = planStr,
                                ["agent_transcript"] = transcript,
                                ["workbook_json"] = workbookStr
                            };
                            var opts = new System.Text.Json.JsonSerializerOptions { WriteIndented = true };
                            File.WriteAllText(exportPath, System.Text.Json.JsonSerializer.Serialize(root, opts));
                            Log($"    -> Consolidated export saved: {exportPath}\r\n");
                        }
                        catch (Exception ex)
                        {
                            Log($"    -> Failed to save consolidated export: {ex.Message}\r\n");
                        }
                    }
                    else
                    {
                        Log($"  - Step {stepNo}: Unsupported action '{step.Action}'.\r\n");
                    }
                }
                catch (Exception ex)
                {
                    Log($"    !! Error: {ex.Message}\r\n");
                }
                finally { stepNo++; }
            }
            Log("> Done.\r\n");
        }

        private void RefreshVisibleList()
        {
            try
            {
                _visibleFiles.Clear();
                _lstTests.Items.Clear();
                int filter = _cmbFilter.SelectedIndex; // 0=All, 1=AI, 2=Non-AI
                foreach (var f in _testFiles)
                {
                    bool hasSpec = GetSpecForPath(f) != null;
                    bool include = filter switch { 1 => hasSpec, 2 => !hasSpec, _ => true };
                    if (!include) continue;
                    _visibleFiles.Add(f);
                    string name = Path.GetFileNameWithoutExtension(f);
                    if (name.EndsWith(".workbook")) name = name[..^".workbook".Length];
                    _lstTests.Items.Add(name);
                }
                _lblStatus.Text = _cmbFilter.SelectedIndex switch
                {
                    1 => $"{_visibleFiles.Count} AI test(s) shown (of {_testFiles.Count})",
                    2 => $"{_visibleFiles.Count} non-AI test(s) shown (of {_testFiles.Count})",
                    _ => $"{_visibleFiles.Count} tests shown"
                };
                if (_lstTests.Items.Count > 0) _lstTests.SelectedIndex = Math.Min(_lstTests.SelectedIndex >= 0 ? _lstTests.SelectedIndex : 0, _lstTests.Items.Count - 1);
            }
            catch { }
        }

        private static string SerializePlan(SpreadsheetApp.Core.AI.AIPlan plan)
        {
            try
            {
                var root = new System.Collections.Generic.Dictionary<string, object?>();
                var list = new System.Collections.Generic.List<object?>();
                foreach (var cmd in plan.Commands)
                {
                    if (cmd is SpreadsheetApp.Core.AI.SetValuesCommand sv)
                    {
                        list.Add(new
                        {
                            type = "set_values",
                            start = new { row = sv.StartRow + 1, col = SpreadsheetApp.Core.CellAddress.ColumnIndexToName(sv.StartCol) },
                            values = sv.Values
                        });
                    }
                    else if (cmd is SpreadsheetApp.Core.AI.SetFormulaCommand sf)
                    {
                        list.Add(new
                        {
                            type = "set_formula",
                            start = new { row = sf.StartRow + 1, col = SpreadsheetApp.Core.CellAddress.ColumnIndexToName(sf.StartCol) },
                            formulas = sf.Formulas
                        });
                    }
                    else if (cmd is SpreadsheetApp.Core.AI.ClearRangeCommand cr)
                    {
                        list.Add(new
                        {
                            type = "clear_range",
                            start = new { row = cr.StartRow + 1, col = SpreadsheetApp.Core.CellAddress.ColumnIndexToName(cr.StartCol) },
                            rows = cr.Rows,
                            cols = cr.Cols
                        });
                    }
                    else if (cmd is SpreadsheetApp.Core.AI.SetTitleCommand st)
                    {
                        list.Add(new
                        {
                            type = "set_title",
                            start = new { row = st.StartRow + 1, col = SpreadsheetApp.Core.CellAddress.ColumnIndexToName(st.StartCol) },
                            rows = st.Rows,
                            cols = st.Cols,
                            text = st.Text
                        });
                    }
                    else if (cmd is SpreadsheetApp.Core.AI.RenameSheetCommand rn)
                    {
                        list.Add(new
                        {
                            type = "rename_sheet",
                            index = rn.Index1,
                            old_name = rn.OldName,
                            new_name = rn.NewName
                        });
                    }
                    else if (cmd is SpreadsheetApp.Core.AI.CreateSheetCommand cs)
                    {
                        list.Add(new { type = "create_sheet", name = cs.Name });
                    }
                    else if (cmd is SpreadsheetApp.Core.AI.SortRangeCommand sr)
                    {
                        list.Add(new
                        {
                            type = "sort_range",
                            start = new { row = sr.StartRow + 1, col = SpreadsheetApp.Core.CellAddress.ColumnIndexToName(sr.StartCol) },
                            rows = sr.Rows,
                            cols = sr.Cols,
                            sort_col = SpreadsheetApp.Core.CellAddress.ColumnIndexToName(sr.SortCol),
                            order = sr.Order,
                            has_header = sr.HasHeader
                        });
                    }
                }
                root["commands"] = list;
                var opts = new System.Text.Json.JsonSerializerOptions { WriteIndented = true };
                return System.Text.Json.JsonSerializer.Serialize(root, opts);
            }
            catch
            {
                return "{\n  \"commands\": []\n}";
            }
        }

        private void WriteDiffToLog(System.Collections.Generic.Dictionary<string, string> before, System.Collections.Generic.Dictionary<string, string> after)
        {
            try
            {
                var changed = new System.Collections.Generic.List<string>();
                var keys = new System.Collections.Generic.HashSet<string>(before.Keys, System.StringComparer.OrdinalIgnoreCase);
                foreach (var k in after.Keys) keys.Add(k);
                foreach (var k in keys)
                {
                    before.TryGetValue(k, out var b);
                    after.TryGetValue(k, out var a);
                    if (string.Equals(b ?? string.Empty, a ?? string.Empty, StringComparison.Ordinal)) continue;
                    if (b == null)
                        changed.Add($"      + {k} = {a}");
                    else if (a == null)
                        changed.Add($"      - {k} (cleared from '{b}')");
                    else
                        changed.Add($"      ~ {k}: '{b}' -> '{a}'");
                }
                if (changed.Count == 0)
                {
                    Log("    -> No cell changes detected.\r\n");
                    return;
                }
                int maxShow = 100;
                Log($"    -> Cell changes ({changed.Count}):\r\n");
                for (int i = 0; i < Math.Min(maxShow, changed.Count); i++)
                    Log(changed[i] + "\r\n");
                if (changed.Count > maxShow) Log($"      ... ({changed.Count - maxShow} more)\r\n");
            }
            catch { }
        }

        // --- Specs support ---
        private sealed class TestSpecs
        {
            public System.Collections.Generic.List<TestSpecItem> Tests { get; set; } = new();
        }
        private sealed class TestSpecItem
        {
            public string File { get; set; } = string.Empty;
            public System.Collections.Generic.List<TestStep> Steps { get; set; } = new();
        }
        private sealed class TestStep
        {
            public string Action { get; set; } = "ai_chat"; // currently only ai_chat
            public string? Prompt { get; set; }
            public string? Location { get; set; } // e.g., "A3" or "A2:B7"
            public bool Apply { get; set; } = true;
            public string? Sheet { get; set; }
        }

        private void LoadSpecs()
        {
            try
            {
                string? testsDir = FindTestsDir();
                if (testsDir == null) return;
                string path = Path.Combine(testsDir, "TEST_SPECS.json");
                if (!File.Exists(path)) return;
                var json = File.ReadAllText(path);
                var opts = new System.Text.Json.JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                _specs = System.Text.Json.JsonSerializer.Deserialize<TestSpecs>(json, opts);
            }
            catch { _specs = null; }
        }

        private TestSpecItem? GetSpecForPath(string file)
        {
            if (_specs == null) return null;
            string name = Path.GetFileName(file);
            foreach (var t in _specs.Tests)
            {
                if (string.Equals(Path.GetFileName(t.File), name, StringComparison.OrdinalIgnoreCase)) return t;
            }
            return null;
        }

        private string DescribeSpecForPath(string file)
        {
            var spec = GetSpecForPath(file);
            if (spec == null || spec.Steps == null || spec.Steps.Count == 0) return "(No automated steps for this test)";
            var sb = new System.Text.StringBuilder();
            for (int i = 0; i < spec.Steps.Count; i++)
            {
                var s = spec.Steps[i];
                sb.AppendLine($"Step {i + 1}: action={s.Action}, sheet={s.Sheet ?? "(current)"}, loc={s.Location ?? "(none)"}, apply={s.Apply}");
                sb.AppendLine($"  prompt: {s.Prompt}");
            }
            return sb.ToString();
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Left || keyData == (Keys.Alt | Keys.Left))
            {
                Navigate(-1);
                return true;
            }
            if (keyData == Keys.Right || keyData == (Keys.Alt | Keys.Right))
            {
                Navigate(1);
                return true;
            }
            if (keyData == Keys.Enter)
            {
                LoadSelected();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
    }
}
