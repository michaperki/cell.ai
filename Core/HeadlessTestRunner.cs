using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading;
using SpreadsheetApp.Core.AI;
using SpreadsheetApp.IO;

namespace SpreadsheetApp.Core
{
    // Minimal headless test runner that mirrors the UI Test Runner outputs.
    public static class HeadlessTestRunner
    {
        private sealed class TestSpecs
        {
            public List<TestSpecItem> Tests { get; set; } = new();
        }
        private sealed class TestSpecItem
        {
            public string File { get; set; } = string.Empty;
            public List<TestStep> Steps { get; set; } = new();
        }
        private sealed class TestStep
        {
            public string Action { get; set; } = "ai_chat"; // ai_chat | ai_agent
            public string? Prompt { get; set; }
            public string? Location { get; set; }
            public bool Apply { get; set; } = true;
            public string? Sheet { get; set; }
            // Stage 2: optional assertions/constraints
            public string[]? AllowedCommands { get; set; }
            public bool? ExpectNoWrites { get; set; }
            public bool? ExpectTranscript { get; set; }
            public Dictionary<string, string>? ExpectedCells { get; set; }
            public int? MaxChangesOutside { get; set; }
            public int? MinChangesTotal { get; set; }
            public bool? Strict { get; set; }
        }

        public static int RunAll(string specsPath, string outputDir, bool reflection = false)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(specsPath)) specsPath = Path.Combine("tests", "TEST_SPECS.json");
                if (!File.Exists(specsPath))
                {
                    try { Console.Error.WriteLine($"TEST_SPECS not found: {specsPath}"); } catch { }
                    return 2;
                }
                var json = File.ReadAllText(specsPath);
                var opts = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                var specs = JsonSerializer.Deserialize<TestSpecs>(json, opts) ?? new TestSpecs();
                if (specs.Tests.Count == 0) { try { Console.Error.WriteLine("No tests in TEST_SPECS.json"); } catch { } return 3; }

                int failed = 0, total = 0;
                int totalSteps = 0, passedSteps = 0;
                var testSummaries = new List<Dictionary<string, object?>>();
                foreach (var t in specs.Tests)
                {
                    string file = t.File;
                    if (!Path.IsPathRooted(file))
                    {
                        var baseDir = Path.GetDirectoryName(specsPath) ?? Environment.CurrentDirectory;
                        file = Path.Combine(baseDir, file);
                    }
                    if (!File.Exists(file))
                    {
                        try { Console.Error.WriteLine($"Missing test file: {file}"); } catch { }
                        failed++; continue;
                    }
                    var summary = RunTestDetailed(file, t.Steps, outputDir, reflection);
                    int code = (int)(summary["exit_code"] ?? 1);
                    total++;
                    if (code != 0) failed++;
                    try { totalSteps += Convert.ToInt32(summary.ContainsKey("steps_total") ? summary["steps_total"] : 0); } catch { }
                    try { passedSteps += Convert.ToInt32(summary.ContainsKey("steps_passed") ? summary["steps_passed"] : 0); } catch { }
                    testSummaries.Add(summary);
                }
                try { Console.WriteLine($"Tests completed: {total - failed} passed, {failed} failed."); } catch { }
                // Write aggregated results summary
                try
                {
                    var root = new Dictionary<string, object?>
                    {
                        ["total_tests"] = total,
                        ["failed_tests"] = failed,
                        ["tests"] = testSummaries
                    };
                    string sumPath = Path.Combine(outputDir, "results_summary.json");
                    TryWriteAllText(sumPath, JsonSerializer.Serialize(root, new JsonSerializerOptions { WriteIndented = true }));
                }
                catch { }
                // Scorecard (Stage 2)
                try
                {
                    var score = new Dictionary<string, object?>
                    {
                        ["total_tests"] = total,
                        ["failed_tests"] = failed,
                        ["total_steps"] = totalSteps,
                        ["passed_steps"] = passedSteps,
                        ["pass_rate"] = totalSteps > 0 ? (double)passedSteps / (double)totalSteps : 0.0
                    };
                    string scorePath = Path.Combine(outputDir, "scorecard.json");
                    TryWriteAllText(scorePath, JsonSerializer.Serialize(score, new JsonSerializerOptions { WriteIndented = true }));
                }
                catch { }

                // Aggregate improvement suggestions across steps (Stage 3)
                try
                {
                    var agg = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                    foreach (var t in testSummaries)
                    {
                        try
                        {
                            if (!t.TryGetValue("steps", out var stepsObj) || stepsObj is not IEnumerable<object> stepsList) continue;
                            foreach (var stepObj in stepsList)
                            {
                                if (stepObj is not Dictionary<string, object?> st) continue;
                                if (st.TryGetValue("assert_failures", out var af) && af is IEnumerable<object> arr)
                                {
                                    foreach (var x in arr)
                                    {
                                        string key = Convert.ToString(x) ?? string.Empty; if (string.IsNullOrWhiteSpace(key)) continue;
                                        agg[key] = agg.TryGetValue(key, out var n) ? n + 1 : 1;
                                    }
                                }
                            }
                        }
                        catch { }
                    }
                    string impPath = Path.Combine(outputDir, "improvement_suggestions.json");
                    TryWriteAllText(impPath, JsonSerializer.Serialize(new Dictionary<string, object?>
                    {
                        ["suggestions"] = agg.OrderByDescending(kv => kv.Value).Select(kv => new { key = kv.Key, count = kv.Value }).ToArray()
                    }, new JsonSerializerOptions { WriteIndented = true }));
                }
                catch { }

                // Append to run history (JSONL)
                try
                {
                    // Collect failure keys from this run
                    var failureKeys = new List<string>();
                    foreach (var t in testSummaries)
                    {
                        try
                        {
                            if (!t.TryGetValue("steps", out var stepsObj) || stepsObj is not IEnumerable<object> stepsList) continue;
                            foreach (var stepObj in stepsList)
                            {
                                if (stepObj is not Dictionary<string, object?> st) continue;
                                if (st.TryGetValue("assert_failures", out var af) && af is IEnumerable<object> arr)
                                    foreach (var x in arr) { var k = Convert.ToString(x); if (!string.IsNullOrWhiteSpace(k)) failureKeys.Add(k!); }
                            }
                        }
                        catch { }
                    }
                    // Detect provider/model from first non-null step
                    string? runProvider = null, runModel = null;
                    foreach (var t in testSummaries)
                    {
                        try
                        {
                            if (!t.TryGetValue("steps", out var stepsObj) || stepsObj is not IEnumerable<object> stepsList) continue;
                            foreach (var stepObj in stepsList)
                            {
                                if (stepObj is not Dictionary<string, object?> st) continue;
                                if (runProvider == null && st.TryGetValue("provider", out var pv) && pv is string ps && !string.IsNullOrWhiteSpace(ps)) runProvider = ps;
                                if (runModel == null && st.TryGetValue("model", out var mv) && mv is string ms && !string.IsNullOrWhiteSpace(ms)) runModel = ms;
                                if (runProvider != null && runModel != null) break;
                            }
                            if (runProvider != null) break;
                        }
                        catch { }
                    }
                    var historyEntry = new Dictionary<string, object?>
                    {
                        ["timestamp"] = DateTime.UtcNow.ToString("o"),
                        ["commit"] = GetGitCommitShort(),
                        ["provider"] = runProvider,
                        ["model"] = runModel,
                        ["spec_file"] = Path.GetFileName(specsPath),
                        ["tests_total"] = total,
                        ["tests_passed"] = total - failed,
                        ["steps_total"] = totalSteps,
                        ["steps_passed"] = passedSteps,
                        ["pass_rate"] = totalSteps > 0 ? Math.Round((double)passedSteps / totalSteps, 3) : 0.0,
                        ["failure_keys"] = failureKeys.Distinct().ToArray()
                    };
                    string historyPath = Path.Combine(outputDir, "run_history.jsonl");
                    string line = JsonSerializer.Serialize(historyEntry);
                    File.AppendAllText(historyPath, line + Environment.NewLine);
                }
                catch { }

                // Generate dashboard
                try { GenerateDashboard(outputDir, specsPath); } catch { }

                return failed > 0 ? 1 : 0;
            }
            catch (Exception ex)
            {
                try { Console.Error.WriteLine($"Headless runner error: {ex}"); } catch { }
                return 1;
            }
        }

        public static int RunSingle(string filePath, string outputDir, bool reflection = false)
        {
            try
            {
                if (!File.Exists(filePath)) { try { Console.Error.WriteLine($"File not found: {filePath}"); } catch { } return 2; }
                // Find steps from tests/TEST_SPECS.json next to the file if present
                string? testsDir = FindTestsDir(Path.GetDirectoryName(filePath) ?? Environment.CurrentDirectory);
                List<TestStep> steps = new();
                if (testsDir != null)
                {
                    string specsPath = Path.Combine(testsDir, "TEST_SPECS.json");
                    if (File.Exists(specsPath))
                    {
                        try
                        {
                            var json = File.ReadAllText(specsPath);
                            var opts = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                            var specs = JsonSerializer.Deserialize<TestSpecs>(json, opts);
                            if (specs != null && specs.Tests != null)
                            {
                                string name = Path.GetFileName(filePath);
                                var item = specs.Tests.FirstOrDefault(x => string.Equals(Path.GetFileName(x.File), name, StringComparison.OrdinalIgnoreCase));
                                if (item != null && item.Steps != null) steps = item.Steps;
                            }
                        }
                        catch { }
                    }
                }
                if (steps.Count == 0)
                {
                    try { Console.Error.WriteLine("No steps found for this file in TEST_SPECS.json"); } catch { }
                    return 4;
                }
                var summary = RunTestDetailed(filePath, steps, outputDir, reflection);
                // Also emit a summary file for single runs
                try { Directory.CreateDirectory(outputDir); File.WriteAllText(Path.Combine(outputDir, "results_summary.json"), JsonSerializer.Serialize(new { total_tests = 1, failed_tests = (int)summary["exit_code"] == 0 ? 0 : 1, tests = new[] { summary } }, new JsonSerializerOptions { WriteIndented = true })); } catch { }
                return (int)(summary["exit_code"] ?? 1);
            }
            catch (Exception ex)
            {
                try { Console.Error.WriteLine($"Headless single error: {ex}"); } catch { }
                return 1;
            }
        }

        private static Dictionary<string, object?> RunTestDetailed(string testFile, List<TestStep> steps, string outputDir, bool reflection)
        {
            try
            {
                Directory.CreateDirectory(outputDir);
                var wb = SpreadsheetIO.LoadWorkbookFromFile(testFile);
                var sheets = wb.Sheets;
                var names = wb.Names;
                int activeSheet = 0;
                if (sheets.Count == 0)
                {
                    try { Console.Error.WriteLine($"Empty workbook: {testFile}"); } catch { }
                    return new Dictionary<string, object?>
                    {
                        ["file"] = Path.GetFileName(testFile),
                        ["steps"] = Array.Empty<object>(),
                        ["exit_code"] = 1,
                        ["passed"] = false
                    };
                }

                string baseName = Path.GetFileNameWithoutExtension(testFile);
                var planner = new ProviderChatPlanner("Auto");
                int stepNo = 1;
                int failures = 0;
                var stepSummaries = new List<Dictionary<string, object?>>();
                foreach (var step in steps)
                {
                    // Activate sheet if provided
                    if (!string.IsNullOrWhiteSpace(step.Sheet))
                    {
                        int idx = names.FindIndex(n => string.Equals(n, step.Sheet, StringComparison.OrdinalIgnoreCase));
                        if (idx >= 0) activeSheet = idx;
                    }
                    var sheet = sheets[activeSheet];
                    // BEFORE snapshot
                    string beforeSnap = Path.Combine(outputDir, $"{baseName}_step{stepNo}.before.workbook.json");
                    TrySaveWorkbook(sheets, names, beforeSnap);
                    var beforeMap = CaptureMap(sheet);

                    // Build context
                    BuildContextForLocation(sheet, names[activeSheet], step.Location, out int sr, out int sc, out int rows, out int cols, out string title, out string[][] selVals);
                    var ctx = new AIContext
                    {
                        SheetName = names[activeSheet],
                        StartRow = sr,
                        StartCol = sc,
                        Rows = rows,
                        Cols = cols,
                        Title = title,
                        SelectionValues = selVals,
                        NearbyValues = BuildNearby(sheet, sr, sc),
                        Workbook = BuildWorkbookSummary(sheets, names),
                        RequestQueriesOnly = false,
                        AnswerOnly = false
                    };
                    // Inject step-level policy into planning context
                    try
                    {
                        // Always fence writes to the selection during headless evaluation
                        ctx.SelectionHardMode = true;
                        // Pass explicit allowlist when provided
                        if (step.AllowedCommands != null)
                        {
                            ctx.AllowedCommands = step.AllowedCommands;
                            // Empty allowlist → answer-only/Q&A mode
                            if (step.AllowedCommands.Length == 0)
                            {
                                ctx.AnswerOnly = true;
                            }
                        }
                    }
                    catch { }

                    AIPlan plan;
                    string[] transcript = Array.Empty<string>();
                    if (string.Equals(step.Action, "ai_agent", StringComparison.OrdinalIgnoreCase))
                    {
                        var res = AgentLoop.RunAsync(planner, sheet, ctx, step.Prompt ?? string.Empty, CancellationToken.None).GetAwaiter().GetResult();
                        plan = res.Plan ?? new AIPlan();
                        transcript = res.Transcript ?? Array.Empty<string>();
                        // Save transcript
                        string logPath = Path.Combine(outputDir, $"{baseName}_step{stepNo}.agent.txt");
                        TryWriteAllText(logPath, transcript.Length > 0 ? string.Join("\n", transcript) : "(no transcript)");
                        // Observations JSON
                        string obsPath = Path.Combine(outputDir, $"{baseName}_step{stepNo}.observations.json");
                        TryWriteAllText(obsPath, JsonSerializer.Serialize(new Dictionary<string, object?> { ["observations"] = transcript }, new JsonSerializerOptions { WriteIndented = true }));
                    }
                    else
                    {
                        plan = planner.PlanAsync(ctx, step.Prompt ?? string.Empty, CancellationToken.None).GetAwaiter().GetResult();
                    }

                    // Plan JSON (prefer RawJson)
                    string planPath = Path.Combine(outputDir, $"{baseName}_step{stepNo}.plan.json");
                    string planJson = string.IsNullOrWhiteSpace(plan.RawJson) ? SerializePlan(plan) : plan.RawJson!;
                    TryWriteAllText(planPath, planJson);

                    // Apply if requested
                    if (step.Apply && plan.Commands != null && plan.Commands.Count > 0)
                    {
                        if (!string.IsNullOrWhiteSpace(step.Location) && step.Location!.Contains(":", StringComparison.Ordinal))
                        {
                            if (CellAddress.TryParse(step.Location.Split(':')[0].Trim(), out int r1, out int c1) && CellAddress.TryParse(step.Location.Split(':')[1].Trim(), out int r2, out int c2))
                            {
                                int rStart = Math.Max(0, Math.Min(r1, r2));
                                int rEnd = Math.Min(sheet.Rows - 1, Math.Max(r1, r2));
                                int cStart = Math.Max(0, Math.Min(c1, c2));
                                int cEnd = Math.Min(sheet.Columns - 1, Math.Max(c1, c2));
                                plan = SanitizePlanToBounds(plan, rStart, cStart, rEnd, cEnd);
                            }
                        }
                        ApplyPlan(plan, sheets, names, ref activeSheet);
                    }

                    // AFTER snapshot and consolidated export
                    string afterSnap = Path.Combine(outputDir, $"{baseName}_step{stepNo}.workbook.json");
                    TrySaveWorkbook(sheets, names, afterSnap);
                    var afterMap = CaptureMap(sheets[activeSheet]);

                    ComputeDiffSummary(beforeMap, afterMap, step.Location, out int totalChanged, out int outsideSelection, out var sample);

                    // Stage 2: assertions
                    var assertFailures = EvaluateAsserts(step, plan, transcript, beforeMap, afterMap, totalChanged, outsideSelection);
                    // Structural operations (insert/delete rows/cols, copy/move range) legitimately affect cells outside the selection.
                    bool penalizeOutside = true;
                    try
                    {
                        if (plan != null && plan.Commands != null && plan.Commands.Count > 0)
                        {
                            foreach (var c in plan.Commands)
                            {
                                if (c is AI.InsertColsCommand || c is AI.DeleteColsCommand || c is AI.InsertRowsCommand || c is AI.DeleteRowsCommand || c is AI.CopyRangeCommand || c is AI.MoveRangeCommand)
                                {
                                    penalizeOutside = false; break;
                                }
                            }
                        }
                    }
                    catch { }
                    bool stepPassed = (penalizeOutside ? outsideSelection == 0 : true) && assertFailures.Count == 0;
                    if (!stepPassed) failures++;

                    string userStr = plan.RawUser ?? $"(no RawUser from provider)\nstepPrompt: {step.Prompt}\nlocation: {step.Location}\n";
                    string sysStr = plan.RawSystem ?? "(no system prompt)";
                    string planStr = string.IsNullOrWhiteSpace(plan.RawJson) ? SerializePlan(plan) : plan.RawJson!;
                    string beforeWb = SafeReadText(beforeSnap);
                    string afterWb = SafeReadText(afterSnap);
                    string planSummary = string.Join("; ", (plan.Commands != null && plan.Commands.Count > 0 ? plan.Commands : Enumerable.Empty<IAICommand>()).Select(c => c.Summarize()));
                    // Provider/usage metadata
                    Dictionary<string, object?>? usageObj = null;
                    try
                    {
                        if (plan?.Usage != null)
                        {
                            usageObj = new Dictionary<string, object?>
                            {
                                ["input_tokens"] = plan.Usage.InputTokens,
                                ["output_tokens"] = plan.Usage.OutputTokens,
                                ["total_tokens"] = plan.Usage.TotalTokens,
                                ["context_limit"] = plan.Usage.ContextLimit,
                                ["remaining_context"] = plan.Usage.RemainingContext
                            };
                        }
                    }
                    catch { }

                    var exportRoot = new Dictionary<string, object?>
                    {
                        ["test_file"] = Path.GetFileName(testFile),
                        ["step"] = stepNo,
                        ["sheet"] = step.Sheet ?? names[activeSheet],
                        ["location"] = step.Location,
                        ["apply"] = step.Apply,
                        ["user"] = userStr,
                        ["system"] = sysStr,
                        ["plan_json"] = planStr,
                        ["plan_summary"] = planSummary,
                        ["provider"] = plan.Provider,
                        ["model"] = plan.Model,
                        ["latency_ms"] = plan.LatencyMs,
                        ["usage"] = usageObj,
                        ["agent_transcript"] = transcript,
                        ["chat_thread"] = Array.Empty<object>(),
                        ["before_workbook_json"] = beforeWb,
                        ["after_workbook_json"] = afterWb,
                        ["changes_total"] = totalChanged,
                        ["changes_outside_selection"] = outsideSelection,
                        ["changes_sample"] = sample,
                        ["assert_failures"] = assertFailures
                    };
                    string exportPath = Path.Combine(outputDir, $"{baseName}_step{stepNo}.export.json");
                    TryWriteAllText(exportPath, JsonSerializer.Serialize(exportRoot, new JsonSerializerOptions { WriteIndented = true }));

                    // Optional: developer reflection
                    if (reflection)
                    {
                        try
                        {
                            var fb = BuildHeuristicReflection(step, plan, totalChanged, outsideSelection, transcript, assertFailures);
                            string fbPath = Path.Combine(outputDir, $"{baseName}_step{stepNo}.feedback.json");
                            TryWriteAllText(fbPath, JsonSerializer.Serialize(fb, new JsonSerializerOptions { WriteIndented = true }));
                        }
                        catch { }
                    }

                    // Per-step summary for aggregated results
                    stepSummaries.Add(new Dictionary<string, object?>
                    {
                        ["index"] = stepNo,
                        ["action"] = step.Action,
                        ["apply"] = step.Apply,
                        ["location"] = step.Location,
                        ["plan_summary"] = planSummary,
                        ["provider"] = plan.Provider,
                        ["model"] = plan.Model,
                        ["latency_ms"] = plan.LatencyMs,
                        ["input_tokens"] = plan.Usage?.InputTokens,
                        ["output_tokens"] = plan.Usage?.OutputTokens,
                        ["total_tokens"] = plan.Usage?.TotalTokens,
                        ["changes_total"] = totalChanged,
                        ["changes_outside_selection"] = outsideSelection,
                        ["assert_failures"] = assertFailures,
                        ["passed"] = stepPassed
                    });

                    stepNo++;
                }
                int stepsPassedCount = 0;
                try { stepsPassedCount = stepSummaries.Count(s => (s.TryGetValue("passed", out var p) && p is bool bp && bp)); } catch { }
                var testSummary = new Dictionary<string, object?>
                {
                    ["file"] = Path.GetFileName(testFile),
                    ["steps"] = stepSummaries,
                    ["steps_total"] = stepSummaries.Count,
                    ["steps_passed"] = stepsPassedCount,
                    ["exit_code"] = failures > 0 ? 1 : 0,
                    ["passed"] = failures == 0
                };
                return testSummary;
            }
            catch (Exception ex)
            {
                try { Console.Error.WriteLine($"RunTest error: {ex}"); } catch { }
                return new Dictionary<string, object?> { ["file"] = Path.GetFileName(testFile), ["steps"] = Array.Empty<object>(), ["exit_code"] = 1, ["passed"] = false };
            }
        }

        // --- Helpers ---
        private static string? FindTestsDir(string start)
        {
            try
            {
                string dir = start;
                for (int i = 0; i < 4; i++)
                {
                    string p = Path.Combine(dir, "tests");
                    if (Directory.Exists(p)) return p;
                    var parent = Directory.GetParent(dir); if (parent == null) break; dir = parent.FullName;
                }
            }
            catch { }
            return null;
        }

        private static void TrySaveWorkbook(IReadOnlyList<Spreadsheet> sheets, IReadOnlyList<string> names, string path)
        {
            try { Directory.CreateDirectory(Path.GetDirectoryName(path) ?? "."); SpreadsheetIO.SaveWorkbookToFile(sheets, names, path); } catch { }
        }
        private static void TryWriteAllText(string path, string contents)
        {
            try { Directory.CreateDirectory(Path.GetDirectoryName(path) ?? "."); File.WriteAllText(path, contents); } catch { }
        }
        private static string SafeReadText(string path)
        {
            try { return File.ReadAllText(path); } catch { return string.Empty; }
        }

        private static Dictionary<string, string> CaptureMap(Spreadsheet sheet)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                for (int r = 0; r < sheet.Rows; r++)
                {
                    for (int c = 0; c < sheet.Columns; c++)
                    {
                        var raw = sheet.GetRaw(r, c);
                        if (!string.IsNullOrWhiteSpace(raw))
                        {
                            string addr = CellAddress.ToAddress(r, c);
                            map[addr] = raw!;
                        }
                    }
                }
            }
            catch { }
            return map;
        }

        private static void BuildContextForLocation(Spreadsheet sheet, string sheetName, string? location, out int sr, out int sc, out int rows, out int cols, out string title, out string[][] selection)
        {
            sr = 0; sc = 0; rows = 20; cols = 6; selection = Array.Empty<string[]>();
            if (!string.IsNullOrWhiteSpace(location))
            {
                string loc = location.Trim();
                if (loc.Contains(":", StringComparison.Ordinal))
                {
                    var parts = loc.Split(':');
                    if (parts.Length == 2 && CellAddress.TryParse(parts[0].Trim(), out int r1, out int c1) && CellAddress.TryParse(parts[1].Trim(), out int r2, out int c2))
                    {
                        int rStart = Math.Max(0, Math.Min(r1, r2));
                        int rEnd = Math.Min(sheet.Rows - 1, Math.Max(r1, r2));
                        int cStart = Math.Max(0, Math.Min(c1, c2));
                        int cEnd = Math.Min(sheet.Columns - 1, Math.Max(c1, c2));
                        sr = rStart; sc = cStart; rows = rEnd - rStart + 1; cols = cEnd - cStart + 1;
                        selection = GetValues(sheet, sr, sc, rows, cols);
                    }
                }
                else if (CellAddress.TryParse(loc, out int rr, out int cc))
                {
                    sr = Math.Max(0, Math.Min(sheet.Rows - 1, rr));
                    sc = Math.Max(0, Math.Min(sheet.Columns - 1, cc));
                }
            }
            title = sr > 0 ? (sheet.GetRaw(sr - 1, sc) ?? string.Empty) : string.Empty;
        }

        private static string[][] GetValues(Spreadsheet sheet, int r0, int c0, int rows, int cols)
        {
            var sel = new string[Math.Max(1, rows)][];
            for (int r = 0; r < rows; r++)
            {
                sel[r] = new string[Math.Max(1, cols)];
                for (int c = 0; c < cols; c++) sel[r][c] = sheet.GetDisplay(r0 + r, c0 + c);
            }
            return sel;
        }

        private static string[][] BuildNearby(Spreadsheet sheet, int sr, int sc)
        {
            try
            {
                int winRows = 20, winCols = 10;
                int startR = Math.Max(0, sr - 1);
                int startC = Math.Max(0, sc - 1);
                int rows = Math.Min(winRows, sheet.Rows - startR);
                int cols = Math.Min(winCols, sheet.Columns - startC);
                return GetValues(sheet, startR, startC, rows, cols);
            }
            catch { return Array.Empty<string[]>(); }
        }

        private static SheetSummary[] BuildWorkbookSummary(IReadOnlyList<Spreadsheet> sheets, IReadOnlyList<string> names)
        {
            var list = new List<SheetSummary>();
            try
            {
                for (int i = 0; i < sheets.Count; i++)
                {
                    var sh = sheets[i];
                    int usedRows = 0, usedCols = 0; string? usedTop = null, usedBottom = null; string[]? header = null; int hdrIdx = -1; int dataCount = 0;
                    for (int r = 0; r < sh.Rows; r++)
                    {
                        bool any = false;
                        for (int c = 0; c < sh.Columns; c++)
                        {
                            var raw = sh.GetRaw(r, c);
                            if (!string.IsNullOrWhiteSpace(raw)) { any = true; usedCols = Math.Max(usedCols, c + 1); }
                        }
                        if (any) usedRows = Math.Max(usedRows, r + 1);
                    }
                    if (usedRows > 0 && usedCols > 0)
                    {
                        usedTop = CellAddress.ToAddress(0, 0);
                        usedBottom = CellAddress.ToAddress(usedRows - 1, usedCols - 1);
                    }
                    hdrIdx = DetectHeaderRow(sh, usedRows, usedCols);
                    if (hdrIdx >= 0 && usedCols > 0)
                    {
                        var hdr = new List<string>();
                        for (int c = 0; c < usedCols; c++) hdr.Add(sh.GetDisplay(hdrIdx, c));
                        header = hdr.ToArray();
                        dataCount = Math.Max(0, usedRows - hdrIdx - 1);
                    }
                    list.Add(new SheetSummary
                    {
                        Name = (names.Count > i ? names[i] : $"Sheet{i + 1}"),
                        UsedRows = usedRows,
                        UsedCols = usedCols,
                        HeaderRow = header,
                        HeaderRowIndex = hdrIdx,
                        DataRowCountExcludingHeader = dataCount,
                        UsedTopLeft = usedTop,
                        UsedBottomRight = usedBottom
                    });
                }
            }
            catch { }
            return list.ToArray();
        }

        private static int DetectHeaderRow(Spreadsheet sh, int usedRows, int usedCols)
        {
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

        private static void ComputeDiffSummary(Dictionary<string, string> before, Dictionary<string, string> after, string? location, out int totalChanged, out int outsideSelection, out List<string> sampleChanges)
        {
            totalChanged = 0; outsideSelection = 0; sampleChanges = new List<string>();
            int rStart = int.MinValue, rEnd = int.MaxValue, cStart = int.MinValue, cEnd = int.MaxValue;
            try
            {
                if (!string.IsNullOrWhiteSpace(location) && location.Contains(":", StringComparison.Ordinal))
                {
                    var parts = location.Split(':');
                    if (parts.Length == 2 && CellAddress.TryParse(parts[0].Trim(), out int rr1, out int cc1) && CellAddress.TryParse(parts[1].Trim(), out int rr2, out int cc2))
                    {
                        rStart = Math.Min(rr1, rr2); rEnd = Math.Max(rr1, rr2);
                        cStart = Math.Min(cc1, cc2); cEnd = Math.Max(cc1, cc2);
                    }
                }
            }
            catch { }
            var keys = new HashSet<string>(before.Keys, StringComparer.OrdinalIgnoreCase);
            foreach (var k in after.Keys) keys.Add(k);
            foreach (var k in keys)
            {
                before.TryGetValue(k, out var b);
                after.TryGetValue(k, out var a);
                if (string.Equals(b ?? string.Empty, a ?? string.Empty, StringComparison.Ordinal)) continue;
                totalChanged++;
                if (CellAddress.TryParse(k, out int rr, out int cc))
                {
                    if (rr < rStart || rr > rEnd || cc < cStart || cc > cEnd) outsideSelection++;
                }
                if (sampleChanges.Count < 20)
                {
                    if (b == null) sampleChanges.Add($"+ {k} = {a}");
                    else if (a == null) sampleChanges.Add($"- {k} (cleared from '{b}')");
                    else sampleChanges.Add($"~ {k}: '{b}' -> '{a}'");
                }
            }
        }

        // Heuristic developer reflection per step
        private static Dictionary<string, object?> BuildHeuristicReflection(TestStep step, AIPlan plan, int changesTotal, int changesOutside, string[] transcript, List<string> assertFailures)
        {
            var missingTools = new List<string>();
            var missingCommands = new List<string>();
            var comms = new List<string>();
            var nextCaps = new List<string>();

            string prompt = step.Prompt?.ToLowerInvariant() ?? string.Empty;
            var cmdTypes = new HashSet<string>(plan?.Commands?.Select(c => c.Type.ToString().ToLowerInvariant()) ?? Enumerable.Empty<string>());

            void Want(string kw, string cmd)
            {
                if (prompt.Contains(kw) && !cmdTypes.Contains(cmd)) missingCommands.Add(cmd);
            }
            Want("insert column", "insertcols");
            Want("delete column", "deletecols");
            Want("insert row", "insertrows");
            Want("delete row", "deleterows");
            Want("rename sheet", "renamesheet");
            Want("sort", "sortrange");
            Want("copy ", "copyrange");
            Want("move ", "moverange");
            Want("format", "setformat");
            Want("validation", "setvalidation");
            if (prompt.Contains("conditional format") || prompt.Contains("highlight")) Want("conditional", "setconditionalformat");
            if (prompt.Contains("normalize") || prompt.Contains("cleanup") || prompt.Contains("transform")) Want("cleanup", "transformrange");

            if ((plan?.Commands?.Count ?? 0) == 0 && step.Apply)
            {
                missingTools.Add("planner_coverage_for_prompt");
                comms.Add("Be explicit about destination (range/sheet) and values vs formulas");
            }
            if (string.IsNullOrWhiteSpace(step.Location))
            {
                comms.Add("Specify a range when writes must stay within bounds");
            }
            if (changesOutside > 0)
            {
                comms.Add("Enable stricter selection fencing (hard mode) or narrow the location");
            }
            if (string.Equals(step.Action, "ai_agent", StringComparison.OrdinalIgnoreCase) && (transcript?.Length ?? 0) == 0)
            {
                missingTools.Add("agent_observation_queries");
                comms.Add("Include a small set of observations before proposing writes");
            }
            // Next capabilities — simple suggestions
            if (!cmdTypes.Contains("setformula") && prompt.Contains("total")) nextCaps.Add("formula_autocomplete_patterns");
            if (!cmdTypes.Contains("transformrangecommand") && prompt.Contains("normalize")) nextCaps.Add("common_text_transforms");
            if (!cmdTypes.Contains("sortrange") && prompt.Contains("sort")) nextCaps.Add("sort_intent_discovery");

            return new Dictionary<string, object?>
            {
                ["missing_tools"] = missingTools.Distinct().ToArray(),
                ["missing_commands"] = missingCommands.Distinct().ToArray(),
                ["comms_feedback"] = comms.Distinct().ToArray(),
                ["next_capabilities"] = nextCaps.Distinct().ToArray(),
                ["assert_failures"] = assertFailures.Distinct().ToArray(),
                ["notes"] = "heuristic"
            };
        }

        // Stage 2: evaluate assertions for a step
        private static List<string> EvaluateAsserts(TestStep step, AIPlan plan, string[] transcript, Dictionary<string, string> beforeMap, Dictionary<string, string> afterMap, int totalChanged, int outsideSelection)
        {
            var failures = new List<string>();
            try
            {
                var cmds = plan?.Commands ?? new List<IAICommand>();
                var allowed = new HashSet<string>((step.AllowedCommands ?? Array.Empty<string>()).Select(s => (s ?? string.Empty).Trim().ToLowerInvariant()));
                if (allowed.Count > 0)
                {
                    foreach (var c in cmds)
                    {
                        var schemaType = CommandTypeToSchema(c.Type);
                        if (!string.IsNullOrEmpty(schemaType) && !allowed.Contains(schemaType)) failures.Add($"unexpected_command:{schemaType}");
                    }
                }
                if ((step.Strict ?? false))
                {
                    bool allowTitle = allowed.Contains("set_title");
                    foreach (var c in cmds)
                    {
                        var schemaType = CommandTypeToSchema(c.Type);
                        if (schemaType == "set_title" && !allowTitle) failures.Add("unexpected_set_title");
                    }
                }
                if (step.ExpectNoWrites == true && totalChanged > 0) failures.Add("expected_no_writes_but_found_changes");
                if (step.ExpectTranscript == true && (transcript == null || transcript.Length == 0)) failures.Add("expected_transcript_missing");
                if (step.MaxChangesOutside.HasValue && outsideSelection > step.MaxChangesOutside.Value)
                    failures.Add($"changes_outside_exceeded:{outsideSelection}>{step.MaxChangesOutside.Value}");
                if (step.MinChangesTotal.HasValue && totalChanged < step.MinChangesTotal.Value)
                    failures.Add($"min_changes_total_not_met:{totalChanged}<{step.MinChangesTotal.Value}");
                if (step.ExpectedCells != null)
                {
                    foreach (var kv in step.ExpectedCells)
                    {
                        string addr = kv.Key?.Trim() ?? string.Empty; if (string.IsNullOrWhiteSpace(addr)) continue;
                        string expected = kv.Value ?? string.Empty;
                        afterMap.TryGetValue(addr, out var got);
                        if (!string.Equals(got ?? string.Empty, expected, StringComparison.Ordinal))
                            failures.Add($"cell_mismatch:{addr} expected='{expected}' got='{(got ?? string.Empty)}'");
                    }
                }
            }
            catch { }
            return failures;
        }

        private static string CommandTypeToSchema(AICommandType t)
        {
            switch (t)
            {
                case AICommandType.SetValues: return "set_values";
                case AICommandType.SetTitle: return "set_title";
                case AICommandType.CreateSheet: return "create_sheet";
                case AICommandType.ClearRange: return "clear_range";
                case AICommandType.RenameSheet: return "rename_sheet";
                case AICommandType.SetFormula: return "set_formula";
                case AICommandType.SortRange: return "sort_range";
                case AICommandType.InsertRows: return "insert_rows";
                case AICommandType.DeleteRows: return "delete_rows";
                case AICommandType.InsertCols: return "insert_cols";
                case AICommandType.DeleteCols: return "delete_cols";
                case AICommandType.DeleteSheet: return "delete_sheet";
                case AICommandType.CopyRange: return "copy_range";
                case AICommandType.MoveRange: return "move_range";
                case AICommandType.SetFormat: return "set_format";
                case AICommandType.SetValidation: return "set_validation";
                case AICommandType.SetConditionalFormat: return "set_conditional_format";
                case AICommandType.TransformRange: return "transform_range";
                default: return t.ToString().ToLowerInvariant();
            }
        }

        private static string SerializePlan(AIPlan plan)
        {
            try
            {
                var root = new Dictionary<string, object?>();
                var list = new List<Dictionary<string, object?>>();
                if (plan.Commands != null)
                {
                    foreach (var cmd in plan.Commands)
                    {
                        var d = new Dictionary<string, object?>();
                        switch (cmd)
                        {
                            case SetValuesCommand sv:
                                d["type"] = "set_values"; d["start"] = new { row = sv.StartRow + 1, col = CellAddress.ColumnIndexToName(sv.StartCol) }; d["values"] = sv.Values; break;
                            case SetFormulaCommand sf:
                                d["type"] = "set_formula"; d["start"] = new { row = sf.StartRow + 1, col = CellAddress.ColumnIndexToName(sf.StartCol) }; d["formulas"] = sf.Formulas; break;
                            case SetTitleCommand st:
                                d["type"] = "set_title"; d["start"] = new { row = st.StartRow + 1, col = CellAddress.ColumnIndexToName(st.StartCol) }; d["rows"] = st.Rows; d["cols"] = st.Cols; d["text"] = st.Text; break;
                            case CreateSheetCommand cs:
                                d["type"] = "create_sheet"; d["name"] = cs.Name; break;
                            case ClearRangeCommand cr:
                                d["type"] = "clear_range"; d["start"] = new { row = cr.StartRow + 1, col = CellAddress.ColumnIndexToName(cr.StartCol) }; d["rows"] = cr.Rows; d["cols"] = cr.Cols; break;
                            case RenameSheetCommand rn:
                                d["type"] = "rename_sheet"; if (rn.Index1.HasValue) d["index"] = rn.Index1.Value; if (!string.IsNullOrWhiteSpace(rn.OldName)) d["old_name"] = rn.OldName; d["new_name"] = rn.NewName; break;
                            case SortRangeCommand sr:
                                d["type"] = "sort_range"; d["start"] = new { row = sr.StartRow + 1, col = CellAddress.ColumnIndexToName(sr.StartCol) }; d["rows"] = sr.Rows; d["cols"] = sr.Cols; d["sort_col"] = CellAddress.ColumnIndexToName(sr.SortCol); d["order"] = sr.Order; d["has_header"] = sr.HasHeader; break;
                            case InsertRowsCommand ir:
                                d["type"] = "insert_rows"; d["at"] = ir.At + 1; d["count"] = ir.Count; break;
                            case DeleteRowsCommand dr:
                                d["type"] = "delete_rows"; d["at"] = dr.At + 1; d["count"] = dr.Count; break;
                            case InsertColsCommand ic:
                                d["type"] = "insert_cols"; d["at"] = CellAddress.ColumnIndexToName(ic.At); d["count"] = ic.Count; break;
                            case DeleteColsCommand dc:
                                d["type"] = "delete_cols"; d["at"] = CellAddress.ColumnIndexToName(dc.At); d["count"] = dc.Count; break;
                            case DeleteSheetCommand ds:
                                d["type"] = "delete_sheet"; if (ds.Index1.HasValue) d["index"] = ds.Index1.Value; if (!string.IsNullOrWhiteSpace(ds.Name)) d["name"] = ds.Name; break;
                            case CopyRangeCommand cr2:
                                d["type"] = "copy_range"; d["start"] = new { row = cr2.StartRow + 1, col = CellAddress.ColumnIndexToName(cr2.StartCol) }; d["rows"] = cr2.Rows; d["cols"] = cr2.Cols; d["dest"] = new { row = cr2.DestRow + 1, col = CellAddress.ColumnIndexToName(cr2.DestCol) }; break;
                            case MoveRangeCommand mr:
                                d["type"] = "move_range"; d["start"] = new { row = mr.StartRow + 1, col = CellAddress.ColumnIndexToName(mr.StartCol) }; d["rows"] = mr.Rows; d["cols"] = mr.Cols; d["dest"] = new { row = mr.DestRow + 1, col = CellAddress.ColumnIndexToName(mr.DestCol) }; break;
                            case SetFormatCommand sfmt:
                                d["type"] = "set_format"; d["start"] = new { row = sfmt.StartRow + 1, col = CellAddress.ColumnIndexToName(sfmt.StartCol) }; d["rows"] = sfmt.Rows; d["cols"] = sfmt.Cols; d["bold"] = sfmt.Bold; d["number_format"] = sfmt.NumberFormat; d["h_align"] = sfmt.HAlign; if (sfmt.ForeColorArgb.HasValue) d["fore_color"] = $"#{sfmt.ForeColorArgb.Value & 0xFFFFFF:X6}"; if (sfmt.BackColorArgb.HasValue) d["back_color"] = $"#{sfmt.BackColorArgb.Value & 0xFFFFFF:X6}"; break;
                            case SetValidationCommand vcmd:
                                d["type"] = "set_validation"; d["start"] = new { row = vcmd.StartRow + 1, col = CellAddress.ColumnIndexToName(vcmd.StartCol) }; d["rows"] = vcmd.Rows; d["cols"] = vcmd.Cols; d["mode"] = vcmd.Mode; d["allow_empty"] = vcmd.AllowEmpty; d["min"] = vcmd.Min; d["max"] = vcmd.Max; d["allowed"] = vcmd.AllowedList; break;
                            case SetConditionalFormatCommand scf:
                                d["type"] = "set_conditional_format"; d["start"] = new { row = scf.StartRow + 1, col = CellAddress.ColumnIndexToName(scf.StartCol) }; d["rows"] = scf.Rows; d["cols"] = scf.Cols; d["op"] = scf.Operator; d["threshold"] = scf.Threshold; d["bold"] = scf.Bold; if (scf.ForeColorArgb.HasValue) d["fore_color"] = $"#{scf.ForeColorArgb.Value & 0xFFFFFF:X6}"; if (scf.BackColorArgb.HasValue) d["back_color"] = $"#{scf.BackColorArgb.Value & 0xFFFFFF:X6}"; d["number_format"] = scf.NumberFormat; d["h_align"] = scf.HAlign; break;
                            case TransformRangeCommand tr:
                                d["type"] = "transform_range"; d["start"] = new { row = tr.StartRow + 1, col = CellAddress.ColumnIndexToName(tr.StartCol) }; d["rows"] = tr.Rows; d["cols"] = tr.Cols; d["op"] = tr.Op; break;
                            default:
                                d["type"] = cmd.Type.ToString(); d["summary"] = cmd.Summarize(); break;
                        }
                        list.Add(d);
                    }
                }
                root["commands"] = list;
                return JsonSerializer.Serialize(root, new JsonSerializerOptions { WriteIndented = true });
            }
            catch { return "{\n  \"commands\": []\n}"; }
        }

        private static AIPlan SanitizePlanToBounds(AIPlan plan, int rStart, int cStart, int rEnd, int cEnd)
        {
            var outPlan = new AIPlan { RawJson = plan.RawJson, RawUser = plan.RawUser, RawSystem = plan.RawSystem, Provider = plan.Provider, Model = plan.Model, Usage = plan.Usage, LatencyMs = plan.LatencyMs };
            foreach (var cmd in plan.Commands)
            {
                switch (cmd)
                {
                    case SetValuesCommand sv:
                    {
                        int rows = sv.Values.Length;
                        int a1 = sv.StartRow; int b1 = sv.StartCol; int a2 = sv.StartRow + rows - 1;
                        int maxCols = 0; for (int r = 0; r < rows; r++) { int len = sv.Values[r]?.Length ?? 0; if (len > maxCols) maxCols = len; }
                        int b2 = (maxCols > 0) ? (sv.StartCol + maxCols - 1) : (sv.StartCol - 1);
                        int rr1 = Math.Max(rStart, a1); int cc1 = Math.Max(cStart, b1);
                        int rr2 = Math.Min(rEnd, a2); int cc2 = Math.Min(cEnd, b2);
                        if (rr1 > rr2 || cc1 > cc2) break;
                        int newCols = cc2 - cc1 + 1;
                        var rowsOut = new List<string[]>();
                        for (int rr = rr1; rr <= rr2; rr++)
                        {
                            int srcRowIndex = rr - a1;
                            if (srcRowIndex < 0 || srcRowIndex >= rows) continue;
                            int rowLen = sv.Values[srcRowIndex]?.Length ?? 0;
                            int srcFirstCol = b1; int srcLastCol = b1 + Math.Max(0, rowLen) - 1;
                            int cS = Math.Max(cc1, srcFirstCol); int cE = Math.Min(cc2, srcLastCol); int overlap = cE - cS + 1;
                            if (overlap <= 0) continue;
                            var line = new string[newCols]; for (int i = 0; i < newCols; i++) line[i] = string.Empty;
                            for (int c = 0; c < overlap; c++)
                            {
                                int destCol = (cS - cc1) + c; int srcCol = (cS - srcFirstCol) + c;
                                line[destCol] = sv.Values[srcRowIndex][srcCol];
                            }
                            rowsOut.Add(line);
                        }
                        if (rowsOut.Count > 0) outPlan.Commands.Add(new SetValuesCommand { StartRow = rr1, StartCol = cc1, Values = rowsOut.ToArray() });
                        break;
                    }
                    case SetFormulaCommand sf:
                    {
                        int rows = sf.Formulas.Length; int cols = rows > 0 ? sf.Formulas[0].Length : 0;
                        int a1 = sf.StartRow; int b1 = sf.StartCol; int a2 = sf.StartRow + rows - 1; int b2 = sf.StartCol + cols - 1;
                        int rr1 = Math.Max(rStart, a1); int cc1 = Math.Max(cStart, b1);
                        int rr2 = Math.Min(rEnd, a2); int cc2 = Math.Min(cEnd, b2);
                        if (rr1 > rr2 || cc1 > cc2) break;
                        int newRows = rr2 - rr1 + 1; int newCols = cc2 - cc1 + 1;
                        var outArr = new string[newRows][];
                        for (int r = 0; r < newRows; r++)
                        {
                            outArr[r] = new string[newCols];
                            for (int c = 0; c < newCols; c++)
                            {
                                int srcR = (rr1 - a1) + r; int srcC = (cc1 - b1) + c;
                                outArr[r][c] = (srcR >= 0 && srcR < rows && srcC >= 0 && srcC < cols) ? (sf.Formulas[srcR][srcC] ?? string.Empty) : string.Empty;
                            }
                        }
                        outPlan.Commands.Add(new SetFormulaCommand { StartRow = rr1, StartCol = cc1, Formulas = outArr });
                        break;
                    }
                    case ClearRangeCommand cr:
                    {
                        int a1 = cr.StartRow; int b1 = cr.StartCol; int a2 = cr.StartRow + Math.Max(1, cr.Rows) - 1; int b2 = cr.StartCol + Math.Max(1, cr.Cols) - 1;
                        int rr1 = Math.Max(rStart, a1); int cc1 = Math.Max(cStart, b1);
                        int rr2 = Math.Min(rEnd, a2); int cc2 = Math.Min(cEnd, b2);
                        if (rr1 > rr2 || cc1 > cc2) break;
                        outPlan.Commands.Add(new ClearRangeCommand { StartRow = rr1, StartCol = cc1, Rows = rr2 - rr1 + 1, Cols = cc2 - cc1 + 1 });
                        break;
                    }
                    case SetTitleCommand st:
                    {
                        int a1 = st.StartRow; int b1 = st.StartCol; int a2 = st.StartRow + Math.Max(1, st.Rows) - 1; int b2 = st.StartCol + Math.Max(1, st.Cols) - 1;
                        int rr1 = Math.Max(rStart, a1); int cc1 = Math.Max(cStart, b1);
                        int rr2 = Math.Min(rEnd, a2); int cc2 = Math.Min(cEnd, b2);
                        if (rr1 > rr2 || cc1 > cc2) break;
                        outPlan.Commands.Add(new SetTitleCommand { StartRow = rr1, StartCol = cc1, Rows = rr2 - rr1 + 1, Cols = cc2 - cc1 + 1, Text = st.Text });
                        break;
                    }
                    default:
                        outPlan.Commands.Add(cmd);
                        break;
                }
            }
            return outPlan;
        }

        private static void ApplyPlan(AIPlan plan, List<Spreadsheet> sheets, List<string> names, ref int activeSheet)
        {
            if (plan == null || plan.Commands == null || plan.Commands.Count == 0) return;
            var sheet = sheets[activeSheet];
            // Apply each command directly to the Spreadsheet model (no UI concerns)
            foreach (var cmd in plan.Commands)
            {
                switch (cmd)
                {
                    case SetValuesCommand set:
                    {
                        int baseRow = set.StartRow;
                        if (baseRow < 0) baseRow = 0;
                        int rows = set.Values.Length;
                        for (int r = 0; r < rows; r++)
                        {
                            int cols = set.Values[r]?.Length ?? 0;
                            for (int c = 0; c < cols; c++)
                            {
                                int rr = baseRow + r; int cc = set.StartCol + c;
                                if (rr < 0 || rr >= sheet.Rows || cc < 0 || cc >= sheet.Columns) continue;
                                string newRaw = set.Values[r][c] ?? string.Empty;
                                if (!sheet.TryValidate(rr, cc, newRaw, out _)) continue;
                                sheet.SetRaw(rr, cc, newRaw);
                                sheet.RecalculateDirty(rr, cc);
                            }
                        }
                        break;
                    }
                    case SetFormulaCommand sf:
                    {
                        int rows = sf.Formulas.Length; int cols = rows > 0 ? sf.Formulas[0].Length : 0;
                        for (int r = 0; r < rows; r++)
                        {
                            for (int c = 0; c < cols; c++)
                            {
                                int rr = sf.StartRow + r; int cc = sf.StartCol + c;
                                if (rr < 0 || rr >= sheet.Rows || cc < 0 || cc >= sheet.Columns) continue;
                                string formula = sf.Formulas[r][c] ?? string.Empty;
                                string newRaw = string.IsNullOrEmpty(formula) ? string.Empty : (formula.StartsWith("=", StringComparison.Ordinal) ? formula : ("=" + formula));
                                sheet.SetRaw(rr, cc, newRaw);
                                sheet.RecalculateDirty(rr, cc);
                            }
                        }
                        break;
                    }
                    case ClearRangeCommand cr:
                    {
                        int rows = Math.Max(1, cr.Rows); int cols = Math.Max(1, cr.Cols);
                        for (int r = 0; r < rows; r++)
                        {
                            for (int c = 0; c < cols; c++)
                            {
                                int rr = cr.StartRow + r; int cc = cr.StartCol + c;
                                if (rr < 0 || rr >= sheet.Rows || cc < 0 || cc >= sheet.Columns) continue;
                                if (!string.IsNullOrEmpty(sheet.GetRaw(rr, cc)))
                                {
                                    sheet.SetRaw(rr, cc, string.Empty); sheet.RecalculateDirty(rr, cc);
                                }
                            }
                        }
                        break;
                    }
                    case RenameSheetCommand rn:
                    {
                        int targetIndex = activeSheet;
                        if (rn.Index1.HasValue) targetIndex = Math.Max(0, Math.Min(sheets.Count - 1, rn.Index1.Value - 1));
                        else if (!string.IsNullOrWhiteSpace(rn.OldName))
                        {
                            int idx = names.FindIndex(n => string.Equals(n, rn.OldName, StringComparison.OrdinalIgnoreCase));
                            if (idx >= 0) targetIndex = idx;
                        }
                        string newName = string.IsNullOrWhiteSpace(rn.NewName) ? names[targetIndex] : rn.NewName.Trim();
                        names[targetIndex] = newName; activeSheet = targetIndex;
                        break;
                    }
                    case SetTitleCommand st:
                    {
                        for (int r = 0; r < Math.Max(1, st.Rows); r++)
                        {
                            for (int c = 0; c < Math.Max(1, st.Cols); c++)
                            {
                                int rr = st.StartRow + r; int cc = st.StartCol + c;
                                if (rr < 0 || rr >= sheet.Rows || cc < 0 || cc >= sheet.Columns) continue;
                                sheet.SetRaw(rr, cc, st.Text);
                                sheet.RecalculateDirty(rr, cc);
                            }
                        }
                        break;
                    }
                    case CreateSheetCommand cs:
                    {
                        string name = string.IsNullOrWhiteSpace(cs.Name) ? $"Sheet{names.Count + 1}" : cs.Name.Trim();
                        var s = new Spreadsheet(Spreadsheet.DefaultRows, Spreadsheet.DefaultCols);
                        sheets.Add(s); names.Add(name); activeSheet = sheets.Count - 1;
                        break;
                    }
                    case SortRangeCommand sr:
                    {
                        int rows = Math.Max(1, sr.Rows); int cols = Math.Max(1, sr.Cols);
                        int r0 = sr.StartRow; int c0 = sr.StartCol; int sortCol = Math.Max(0, Math.Min(sheet.Columns - 1, sr.SortCol));
                        int rStart = r0; int rCount = rows; if (sr.HasHeader && rCount > 1) { rStart = r0 + 1; rCount = rows - 1; }
                        var block = new string[rows][]; for (int r = 0; r < rows; r++) { block[r] = new string[cols]; for (int c = 0; c < cols; c++) block[r][c] = sheet.GetRaw(r0 + r, c0 + c) ?? string.Empty; }
                        var idx = Enumerable.Range(0, rCount).ToList(); int sortRelCol = Math.Max(0, Math.Min(cols - 1, sortCol - c0));
                        idx.Sort((a, b) =>
                        {
                            int ra = rStart - r0 + a; int rb = rStart - r0 + b;
                            var va = sheet.GetValue(r0 + ra, c0 + sortRelCol); var vb = sheet.GetValue(r0 + rb, c0 + sortRelCol);
                            int cmp; if (va.Error == null && vb.Error == null && va.Number is double na && vb.Number is double nb) cmp = na.CompareTo(nb);
                            else { var sa = va.ToDisplay(); var sb = vb.ToDisplay(); bool ea = string.IsNullOrWhiteSpace(sa), eb = string.IsNullOrWhiteSpace(sb); if (ea && eb) cmp = 0; else if (ea) cmp = 1; else if (eb) cmp = -1; else cmp = string.Compare(sa, sb, StringComparison.OrdinalIgnoreCase); }
                            if (string.Equals(sr.Order, "desc", StringComparison.OrdinalIgnoreCase)) cmp = -cmp; return cmp;
                        });
                        var sortedBlock = new string[rows][]; for (int r = 0; r < rows; r++) sortedBlock[r] = new string[cols];
                        if (sr.HasHeader && rows > 0) Array.Copy(block[0], sortedBlock[0], cols);
                        for (int i = 0; i < rCount; i++) { int srcRel = rStart - r0 + idx[i]; int dstRel = rStart - r0 + i; Array.Copy(block[srcRel], sortedBlock[dstRel], cols); }
                        for (int r = 0; r < rows; r++) for (int c = 0; c < cols; c++) { int rr = r0 + r, cc = c0 + c; string newRaw = sortedBlock[r][c] ?? string.Empty; sheet.SetRaw(rr, cc, newRaw); sheet.RecalculateDirty(rr, cc); }
                        break;
                    }
                    case InsertRowsCommand ir:
                    {
                        int at = Math.Max(0, Math.Min(sheet.Rows - 1, ir.At)); int count = Math.Max(1, Math.Min(sheet.Rows - at, ir.Count));
                        for (int r = sheet.Rows - 1; r >= at + count; r--)
                        {
                            int srcRow = r - count;
                            for (int c = 0; c < sheet.Columns; c++) { string? src = sheet.GetRaw(srcRow, c); sheet.SetRaw(r, c, src); }
                        }
                        for (int r = at; r < at + count && r < sheet.Rows; r++) for (int c = 0; c < sheet.Columns; c++) sheet.SetRaw(r, c, string.Empty);
                        sheet.Recalculate();
                        break;
                    }
                    case DeleteRowsCommand dr:
                    {
                        int at = Math.Max(0, Math.Min(sheet.Rows - 1, dr.At)); int count = Math.Max(1, Math.Min(sheet.Rows - at, dr.Count));
                        for (int r = at; r < sheet.Rows - count; r++)
                        {
                            int srcRow = r + count; for (int c = 0; c < sheet.Columns; c++) { string? src = sheet.GetRaw(srcRow, c); sheet.SetRaw(r, c, src); }
                        }
                        for (int r = sheet.Rows - count; r < sheet.Rows; r++) for (int c = 0; c < sheet.Columns; c++) sheet.SetRaw(r, c, string.Empty);
                        sheet.Recalculate();
                        break;
                    }
                    case InsertColsCommand ic:
                    {
                        int at = Math.Max(0, Math.Min(sheet.Columns - 1, ic.At)); int count = Math.Max(1, Math.Min(sheet.Columns - at, ic.Count));
                        for (int c = sheet.Columns - 1; c >= at + count; c--)
                        { int srcCol = c - count; for (int r = 0; r < sheet.Rows; r++) { string? src = sheet.GetRaw(r, srcCol); sheet.SetRaw(r, c, src); } }
                        for (int c = at; c < at + count && c < sheet.Columns; c++) for (int r = 0; r < sheet.Rows; r++) sheet.SetRaw(r, c, string.Empty);
                        sheet.Recalculate();
                        break;
                    }
                    case DeleteColsCommand dc:
                    {
                        int at = Math.Max(0, Math.Min(sheet.Columns - 1, dc.At)); int count = Math.Max(1, Math.Min(sheet.Columns - at, dc.Count));
                        for (int c = at; c < sheet.Columns - count; c++)
                        { int srcCol = c + count; for (int r = 0; r < sheet.Rows; r++) { string? src = sheet.GetRaw(r, srcCol); sheet.SetRaw(r, c, src); } }
                        for (int c = sheet.Columns - count; c < sheet.Columns; c++) for (int r = 0; r < sheet.Rows; r++) sheet.SetRaw(r, c, string.Empty);
                        sheet.Recalculate();
                        break;
                    }
                    case DeleteSheetCommand ds:
                    {
                        int targetIndex = activeSheet;
                        if (ds.Index1.HasValue) targetIndex = Math.Max(0, Math.Min(sheets.Count - 1, ds.Index1.Value - 1));
                        else if (!string.IsNullOrWhiteSpace(ds.Name)) { int idx = names.FindIndex(n => string.Equals(n, ds.Name, StringComparison.OrdinalIgnoreCase)); if (idx >= 0) targetIndex = idx; }
                        if (sheets.Count > 1) { sheets.RemoveAt(targetIndex); names.RemoveAt(targetIndex); activeSheet = Math.Max(0, Math.Min(activeSheet, sheets.Count - 1)); }
                        break;
                    }
                    case CopyRangeCommand copy:
                    {
                        int r0 = Math.Max(0, copy.StartRow); int c0 = Math.Max(0, copy.StartCol);
                        int rows = Math.Max(1, copy.Rows); int cols = Math.Max(1, copy.Cols);
                        int rd = Math.Max(0, copy.DestRow); int cd = Math.Max(0, copy.DestCol);
                        int dRow = rd - r0; int dCol = cd - c0;
                        var raw = new string[rows][];
                        for (int r = 0; r < rows; r++) { raw[r] = new string[cols]; for (int c = 0; c < cols; c++) raw[r][c] = sheet.GetRaw(r0 + r, c0 + c) ?? string.Empty; }
                        for (int r = 0; r < rows; r++)
                        {
                            for (int c = 0; c < cols; c++)
                            {
                                int tr = rd + r, tc = cd + c; if (tr < 0 || tr >= sheet.Rows || tc < 0 || tc >= sheet.Columns) continue;
                                string src = raw[r][c] ?? string.Empty; string newRaw = src;
                                if (!string.IsNullOrEmpty(src) && src.StartsWith("=", StringComparison.Ordinal)) { try { newRaw = RewriteFormulaForPaste(src, dRow, dCol); } catch { newRaw = src; } }
                                sheet.SetRaw(tr, tc, newRaw); sheet.RecalculateDirty(tr, tc);
                            }
                        }
                        break;
                    }
                    case MoveRangeCommand move:
                    {
                        int r0 = Math.Max(0, move.StartRow); int c0 = Math.Max(0, move.StartCol);
                        int rows = Math.Max(1, move.Rows); int cols = Math.Max(1, move.Cols);
                        int rd = Math.Max(0, move.DestRow); int cd = Math.Max(0, move.DestCol);
                        int dRow = rd - r0; int dCol = cd - c0;
                        var raw = new string[rows][];
                        for (int r = 0; r < rows; r++) { raw[r] = new string[cols]; for (int c = 0; c < cols; c++) raw[r][c] = sheet.GetRaw(r0 + r, c0 + c) ?? string.Empty; }
                        for (int r = 0; r < rows; r++)
                        {
                            for (int c = 0; c < cols; c++)
                            {
                                int tr = rd + r, tc = cd + c; if (tr < 0 || tr >= sheet.Rows || tc < 0 || tc >= sheet.Columns) continue;
                                string src = raw[r][c] ?? string.Empty; string newRaw = src;
                                if (!string.IsNullOrEmpty(src) && src.StartsWith("=", StringComparison.Ordinal)) { try { newRaw = RewriteFormulaForPaste(src, dRow, dCol); } catch { newRaw = src; } }
                                sheet.SetRaw(tr, tc, newRaw); sheet.RecalculateDirty(tr, tc);
                            }
                        }
                        // Clear source
                        for (int r = 0; r < rows; r++) for (int c = 0; c < cols; c++) { int srcR = r0 + r, srcC = c0 + c; if (!string.IsNullOrEmpty(sheet.GetRaw(srcR, srcC))) { sheet.SetRaw(srcR, srcC, string.Empty); sheet.RecalculateDirty(srcR, srcC); } }
                        break;
                    }
                    case SetFormatCommand fmt:
                    {
                        int r0 = Math.Max(0, fmt.StartRow); int c0 = Math.Max(0, fmt.StartCol);
                        int rows = Math.Max(1, fmt.Rows); int cols = Math.Max(1, fmt.Cols);
                        for (int r = 0; r < rows; r++) for (int c = 0; c < cols; c++) { int rr = r0 + r, cc = c0 + c; if (rr < 0 || rr >= sheet.Rows || cc < 0 || cc >= sheet.Columns) continue; var cur = sheet.GetFormat(rr, cc) ?? new CellFormat(); if (fmt.Bold.HasValue) cur.Bold = fmt.Bold.Value; if (fmt.ForeColorArgb.HasValue) cur.ForeColorArgb = fmt.ForeColorArgb.Value; if (fmt.BackColorArgb.HasValue) cur.BackColorArgb = fmt.BackColorArgb.Value; if (!string.IsNullOrWhiteSpace(fmt.NumberFormat)) cur.NumberFormat = fmt.NumberFormat; if (!string.IsNullOrWhiteSpace(fmt.HAlign)) { var s = fmt.HAlign!.Trim().ToLowerInvariant(); cur.HAlign = s == "center" ? CellHAlign.Center : s == "right" ? CellHAlign.Right : CellHAlign.Left; } sheet.SetFormat(rr, cc, cur); }
                        break;
                    }
                    case SetValidationCommand vcmd:
                    {
                        int r0 = Math.Max(0, vcmd.StartRow); int c0 = Math.Max(0, vcmd.StartCol);
                        int rows = Math.Max(1, vcmd.Rows); int cols = Math.Max(1, vcmd.Cols);
                        var mode = (vcmd.Mode ?? "none").Trim().ToLowerInvariant();
                        for (int r = 0; r < rows; r++)
                        {
                            for (int c = 0; c < cols; c++)
                            {
                                int rr = r0 + r, cc = c0 + c; if (rr < 0 || rr >= sheet.Rows || cc < 0 || cc >= sheet.Columns) continue;
                                ValidationRule? rule = null;
                                if (mode == "list") rule = new ValidationRule { Mode = ValidationMode.List, AllowEmpty = vcmd.AllowEmpty, AllowedList = vcmd.AllowedList };
                                else if (mode == "number_between") rule = new ValidationRule { Mode = ValidationMode.NumberBetween, AllowEmpty = vcmd.AllowEmpty, Min = vcmd.Min, Max = vcmd.Max };
                                sheet.SetValidation(rr, cc, rule);
                            }
                        }
                        break;
                    }
                    case SetConditionalFormatCommand cfmt:
                    {
                        int r0 = Math.Max(0, cfmt.StartRow); int c0 = Math.Max(0, cfmt.StartCol);
                        int rows = Math.Max(1, cfmt.Rows); int cols = Math.Max(1, cfmt.Cols);
                        string op = (cfmt.Operator ?? ">").Trim(); double th = cfmt.Threshold;
                        for (int r = 0; r < rows; r++)
                        {
                            for (int c = 0; c < cols; c++)
                            {
                                int rr = r0 + r, cc = c0 + c; if (rr < 0 || rr >= sheet.Rows || cc < 0 || cc >= sheet.Columns) continue;
                                var val = sheet.GetValue(rr, cc); bool match = false;
                                if (val.Error == null && val.Number is double d)
                                {
                                    match = op switch { ">" => d > th, ">=" => d >= th, "<" => d < th, "<=" => d <= th, "==" => d == th, "!=" => d != th, _ => false };
                                }
                                if (match)
                                {
                                    var cur = sheet.GetFormat(rr, cc) ?? new CellFormat();
                                    if (cfmt.Bold.HasValue) cur.Bold = cfmt.Bold.Value; if (cfmt.ForeColorArgb.HasValue) cur.ForeColorArgb = cfmt.ForeColorArgb.Value; if (cfmt.BackColorArgb.HasValue) cur.BackColorArgb = cfmt.BackColorArgb.Value; if (!string.IsNullOrWhiteSpace(cfmt.NumberFormat)) cur.NumberFormat = cfmt.NumberFormat; if (!string.IsNullOrWhiteSpace(cfmt.HAlign)) { var s = cfmt.HAlign!.Trim().ToLowerInvariant(); cur.HAlign = s == "center" ? CellHAlign.Center : s == "right" ? CellHAlign.Right : CellHAlign.Left; } sheet.SetFormat(rr, cc, cur);
                                }
                            }
                        }
                        break;
                    }
                    case TransformRangeCommand tr:
                    {
                        int r0 = Math.Max(0, tr.StartRow); int c0 = Math.Max(0, tr.StartCol);
                        int rows = Math.Max(1, tr.Rows); int cols = Math.Max(1, tr.Cols);
                        string op = (tr.Op ?? "trim").Trim().ToLowerInvariant();
                        for (int r = 0; r < rows; r++)
                        {
                            for (int c = 0; c < cols; c++)
                            {
                                int rr = r0 + r, cc = c0 + c; if (rr < 0 || rr >= sheet.Rows || cc < 0 || cc >= sheet.Columns) continue;
                                string oldRaw = sheet.GetRaw(rr, cc) ?? string.Empty; if (string.IsNullOrEmpty(oldRaw)) continue; if (oldRaw.StartsWith("=", StringComparison.Ordinal)) continue;
                                string newRaw = ApplyTransform(oldRaw, op); if (!string.Equals(oldRaw, newRaw, StringComparison.Ordinal)) { sheet.SetRaw(rr, cc, newRaw); sheet.RecalculateDirty(rr, cc); }
                            }
                        }
                        break;
                    }
                }
            }
        }

        private static string ApplyTransform(string text, string op)
        {
            static string CollapseSpaces(string s)
            {
                var parts = s.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                return string.Join(" ", parts);
            }
            switch (op)
            {
                case "trim": return CollapseSpaces(text.Trim());
                case "upper":
                case "uppercase": return text.ToUpperInvariant();
                case "lower":
                case "lowercase": return text.ToLowerInvariant();
                case "proper":
                case "title":
                case "title_case":
                    {
                        var lower = text.ToLowerInvariant();
                        try { return System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(CollapseSpaces(lower).Trim()); }
                        catch { return CollapseSpaces(lower).Trim(); }
                    }
                case "strip_punct":
                    {
                        var arr = text.Where(ch => !char.IsPunctuation(ch)).ToArray();
                        return CollapseSpaces(new string(arr).Trim());
                    }
                case "normalize_city":
                    {
                        var lower = text.Trim().ToLowerInvariant(); lower = CollapseSpaces(lower);
                        try { return System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(lower); }
                        catch { return lower; }
                    }
                default: return text;
            }
        }

        private static string RewriteFormulaForPaste(string src, int dRow, int dCol)
        {
            if (string.IsNullOrEmpty(src) || !src.StartsWith("=", StringComparison.Ordinal)) return src;
            // Simple adjustment: shift A1 references by dRow/dCol via FormulaEngine helper
            // We do not have a direct helper, approximate by scanning A1 refs
            var expr = src.Substring(1);
            var sb = new StringBuilder();
            int i = 0;
            while (i < expr.Length)
            {
                if (TryParseCellToken(expr, ref i, out int r, out int c))
                {
                    int rr = r + dRow; int cc = c + dCol; sb.Append(CellAddress.ToAddress(rr, cc));
                }
                else { sb.Append(expr[i]); i++; }
            }
            return "=" + sb.ToString();
        }

        private static bool TryParseCellToken(string s, ref int i, out int row, out int col)
        {
            row = col = 0; int pos = i;
            if (pos < s.Length && s[pos] == '$') pos++;
            int lettersStart = pos; while (pos < s.Length && char.IsLetter(s[pos])) pos++; if (pos == lettersStart) return false;
            if (pos < s.Length && s[pos] == '$') pos++;
            int digitsStart = pos; while (pos < s.Length && char.IsDigit(s[pos])) pos++; if (pos == digitsStart) return false;
            string letters = s.Substring(lettersStart, pos - lettersStart).Replace("$", string.Empty);
            string digits = s.Substring(digitsStart, pos - digitsStart);
            string addr = (letters + digits).ToUpperInvariant();
            if (!CellAddress.TryParse(addr, out row, out col)) return false; i = pos; return true;
        }

        // --- Run history & dashboard ---

        private static string? GetGitCommitShort()
        {
            try
            {
                var psi = new ProcessStartInfo("git", "rev-parse --short HEAD")
                {
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                using var proc = Process.Start(psi);
                if (proc == null) return null;
                string output = proc.StandardOutput.ReadToEnd().Trim();
                proc.WaitForExit(3000);
                return string.IsNullOrWhiteSpace(output) ? null : output;
            }
            catch { return null; }
        }

        private static void GenerateDashboard(string outputDir, string specsPath)
        {
            string historyPath = Path.Combine(outputDir, "run_history.jsonl");
            if (!File.Exists(historyPath)) return;

            var runs = new List<Dictionary<string, JsonElement>>();
            foreach (var line in File.ReadAllLines(historyPath))
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                try { var doc = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(line); if (doc != null) runs.Add(doc); } catch { }
            }
            if (runs.Count == 0) return;

            // Load capability map if available
            var capMap = new Dictionary<string, List<string>>();
            try
            {
                string? testsDir = Path.GetDirectoryName(specsPath);
                if (testsDir == null) testsDir = FindTestsDir(Environment.CurrentDirectory);
                if (testsDir != null)
                {
                    string capPath = Path.Combine(testsDir, "capability_map.json");
                    if (File.Exists(capPath))
                    {
                        var raw = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(File.ReadAllText(capPath));
                        if (raw != null)
                        {
                            foreach (var kv in raw)
                            {
                                var ids = new List<string>();
                                if (kv.Value.ValueKind == JsonValueKind.Array)
                                    foreach (var el in kv.Value.EnumerateArray()) ids.Add(el.GetString() ?? "");
                                capMap[kv.Key] = ids;
                            }
                        }
                    }
                }
            }
            catch { }

            // Load latest results_summary for per-test pass/fail
            var testResults = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            try
            {
                string sumPath = Path.Combine(outputDir, "results_summary.json");
                if (File.Exists(sumPath))
                {
                    using var doc = JsonDocument.Parse(File.ReadAllText(sumPath));
                    if (doc.RootElement.TryGetProperty("tests", out var testsEl) && testsEl.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var t in testsEl.EnumerateArray())
                        {
                            string? file = t.TryGetProperty("file", out var f) ? f.GetString() : null;
                            bool passed = t.TryGetProperty("passed", out var p) && p.ValueKind == JsonValueKind.True;
                            if (file != null)
                            {
                                // Extract test number from filename like "test_01_..."
                                var match = System.Text.RegularExpressions.Regex.Match(file, @"test_(\d+)");
                                if (match.Success) testResults[match.Groups[1].Value] = passed;
                            }
                        }
                    }
                }
            }
            catch { }

            var sb = new StringBuilder();
            sb.AppendLine("# Dashboard");
            sb.AppendLine();
            sb.AppendLine($"Generated: {DateTime.UtcNow:yyyy-MM-dd HH:mm} UTC");
            sb.AppendLine();

            // Section 1: Latest run summary
            var latest = runs[runs.Count - 1];
            sb.AppendLine("## Latest Run");
            sb.AppendLine();
            sb.AppendLine($"| Metric | Value |");
            sb.AppendLine($"|--------|-------|");
            sb.AppendLine($"| Spec file | {SafeGetString(latest, "spec_file")} |");
            sb.AppendLine($"| Commit | `{SafeGetString(latest, "commit") ?? "unknown"}` |");
            sb.AppendLine($"| Provider | {SafeGetString(latest, "provider") ?? "mock"} |");
            sb.AppendLine($"| Model | {SafeGetString(latest, "model") ?? "—"} |");
            sb.AppendLine($"| Tests | {SafeGetInt(latest, "tests_passed")}/{SafeGetInt(latest, "tests_total")} passed |");
            sb.AppendLine($"| Steps | {SafeGetInt(latest, "steps_passed")}/{SafeGetInt(latest, "steps_total")} passed |");
            sb.AppendLine($"| Pass rate | {SafeGetDouble(latest, "pass_rate"):P1} |");
            var latestFailures = SafeGetStringArray(latest, "failure_keys");
            sb.AppendLine($"| Failure keys | {(latestFailures.Length > 0 ? string.Join(", ", latestFailures.Select(k => $"`{k}`")) : "none")} |");
            sb.AppendLine();

            // Section 2: Trend (last N runs)
            sb.AppendLine("## Run History (last 20)");
            sb.AppendLine();
            sb.AppendLine("| # | Timestamp | Commit | Provider | Spec | Steps | Pass rate | Failures |");
            sb.AppendLine("|---|-----------|--------|----------|------|-------|-----------|----------|");
            int startIdx = Math.Max(0, runs.Count - 20);
            for (int i = runs.Count - 1; i >= startIdx; i--)
            {
                var r = runs[i];
                string ts = SafeGetString(r, "timestamp") ?? "";
                if (ts.Length > 16) ts = ts.Substring(0, 16); // trim to minute
                var fk = SafeGetStringArray(r, "failure_keys");
                sb.AppendLine($"| {i + 1} | {ts} | `{SafeGetString(r, "commit") ?? "?"}` | {SafeGetString(r, "provider") ?? "mock"} | {SafeGetString(r, "spec_file")} | {SafeGetInt(r, "steps_passed")}/{SafeGetInt(r, "steps_total")} | {SafeGetDouble(r, "pass_rate"):P1} | {(fk.Length > 0 ? string.Join(", ", fk.Take(3).Select(k => $"`{k}`")) : "—")} |");
            }
            sb.AppendLine();

            // Section 3: Capability matrix (if map available and results available)
            if (capMap.Count > 0)
            {
                sb.AppendLine("## Capability Matrix");
                sb.AppendLine();
                sb.AppendLine("| Capability | Tests | Covered | Passing | Status |");
                sb.AppendLine("|-----------|-------|---------|---------|--------|");
                foreach (var kv in capMap.OrderBy(x => x.Key))
                {
                    var ids = kv.Value.Where(id => !string.IsNullOrWhiteSpace(id)).ToList();
                    int covered = ids.Count;
                    int passing = 0, failing = 0, unknown = 0;
                    foreach (var id in ids)
                    {
                        if (testResults.TryGetValue(id, out var p)) { if (p) passing++; else failing++; }
                        else unknown++;
                    }
                    string status = covered == 0 ? "no tests" : failing > 0 ? $"{failing} failing" : unknown > 0 ? $"{passing} pass, {unknown} no data" : $"{passing}/{covered} pass";
                    string emoji = covered == 0 ? "—" : failing > 0 ? "FAIL" : unknown > 0 ? "PARTIAL" : "PASS";
                    string name = kv.Key.Replace("_", " ");
                    sb.AppendLine($"| {name} | {string.Join(", ", ids)} | {covered} | {passing} | {emoji}: {status} |");
                }
                sb.AppendLine();
            }

            // Section 4: Aggregated failure patterns across all runs
            sb.AppendLine("## Failure Patterns (all runs)");
            sb.AppendLine();
            var allFailures = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var r in runs)
            {
                var fk = SafeGetStringArray(r, "failure_keys");
                foreach (var k in fk) allFailures[k] = allFailures.TryGetValue(k, out var n) ? n + 1 : 1;
            }
            if (allFailures.Count == 0)
            {
                sb.AppendLine("No failures recorded across any run.");
            }
            else
            {
                sb.AppendLine("| Failure key | Occurrences | First seen | Last seen |");
                sb.AppendLine("|------------|-------------|------------|-----------|");
                foreach (var kv in allFailures.OrderByDescending(x => x.Value))
                {
                    string? first = null, last = null;
                    foreach (var r in runs)
                    {
                        var fk = SafeGetStringArray(r, "failure_keys");
                        if (fk.Contains(kv.Key, StringComparer.OrdinalIgnoreCase))
                        {
                            string ts = SafeGetString(r, "timestamp") ?? "";
                            if (ts.Length > 10) ts = ts.Substring(0, 10);
                            if (first == null) first = ts;
                            last = ts;
                        }
                    }
                    sb.AppendLine($"| `{kv.Key}` | {kv.Value} | {first} | {last} |");
                }
            }
            sb.AppendLine();

            // Section 5: Regressions / improvements (compare last two runs)
            if (runs.Count >= 2)
            {
                var prev = runs[runs.Count - 2];
                var curr = runs[runs.Count - 1];
                double prevRate = SafeGetDouble(prev, "pass_rate");
                double currRate = SafeGetDouble(curr, "pass_rate");
                var prevFk = new HashSet<string>(SafeGetStringArray(prev, "failure_keys"), StringComparer.OrdinalIgnoreCase);
                var currFk = new HashSet<string>(SafeGetStringArray(curr, "failure_keys"), StringComparer.OrdinalIgnoreCase);
                var newFailures = currFk.Except(prevFk).ToList();
                var fixedFailures = prevFk.Except(currFk).ToList();

                sb.AppendLine("## Changes Since Previous Run");
                sb.AppendLine();
                if (currRate > prevRate) sb.AppendLine($"Pass rate: {prevRate:P1} -> {currRate:P1} (improved)");
                else if (currRate < prevRate) sb.AppendLine($"Pass rate: {prevRate:P1} -> {currRate:P1} (REGRESSED)");
                else sb.AppendLine($"Pass rate: {currRate:P1} (unchanged)");
                sb.AppendLine();
                if (fixedFailures.Count > 0) sb.AppendLine($"Fixed: {string.Join(", ", fixedFailures.Select(k => $"`{k}`"))}");
                if (newFailures.Count > 0) sb.AppendLine($"New failures: {string.Join(", ", newFailures.Select(k => $"`{k}`"))}");
                if (fixedFailures.Count == 0 && newFailures.Count == 0) sb.AppendLine("No change in failure patterns.");
                sb.AppendLine();
            }

            // Section 6: What to work on next
            sb.AppendLine("## Suggested Next Actions");
            sb.AppendLine();
            if (capMap.Count > 0)
            {
                var uncovered = capMap.Where(kv => kv.Value.Count == 0).Select(kv => kv.Key.Replace("_", " ")).ToList();
                if (uncovered.Count > 0) sb.AppendLine($"- **Add tests for:** {string.Join(", ", uncovered)}");
                var failingCaps = capMap.Where(kv => kv.Value.Any(id => testResults.TryGetValue(id, out var p) && !p)).Select(kv => kv.Key.Replace("_", " ")).ToList();
                if (failingCaps.Count > 0) sb.AppendLine($"- **Fix failing capabilities:** {string.Join(", ", failingCaps)}");
            }
            if (allFailures.Count > 0)
            {
                var top = allFailures.OrderByDescending(x => x.Value).Take(3).Select(kv => $"`{kv.Key}` ({kv.Value}x)");
                sb.AppendLine($"- **Top recurring failures:** {string.Join(", ", top)}");
            }
            sb.AppendLine();

            string dashPath = Path.Combine(outputDir, "DASHBOARD.md");
            TryWriteAllText(dashPath, sb.ToString());

            // Also emit a .workbook.json that can be opened in the app
            try { GenerateDashboardWorkbook(outputDir, runs, capMap, testResults, allFailures); } catch { }
        }

        private static void GenerateDashboardWorkbook(
            string outputDir,
            List<Dictionary<string, JsonElement>> runs,
            Dictionary<string, List<string>> capMap,
            Dictionary<string, bool> testResults,
            Dictionary<string, int> allFailures)
        {
            // Build workbook JSON directly (avoid dependency on SpreadsheetIO internals)
            var sheets = new List<object>();

            // --- Sheet 1: Overview ---
            {
                var cells = new Dictionary<string, string>();
                var latest = runs[runs.Count - 1];
                cells["A1"] = "Dashboard";
                cells["A2"] = $"Generated: {DateTime.UtcNow:yyyy-MM-dd HH:mm} UTC";
                cells["A4"] = "Metric"; cells["B4"] = "Value";
                cells["A5"] = "Spec file"; cells["B5"] = SafeGetString(latest, "spec_file") ?? "";
                cells["A6"] = "Commit"; cells["B6"] = SafeGetString(latest, "commit") ?? "unknown";
                cells["A7"] = "Provider"; cells["B7"] = SafeGetString(latest, "provider") ?? "mock";
                cells["A8"] = "Model"; cells["B8"] = SafeGetString(latest, "model") ?? "";
                cells["A9"] = "Tests passed"; cells["B9"] = $"{SafeGetInt(latest, "tests_passed")}/{SafeGetInt(latest, "tests_total")}";
                cells["A10"] = "Steps passed"; cells["B10"] = $"{SafeGetInt(latest, "steps_passed")}/{SafeGetInt(latest, "steps_total")}";
                cells["A11"] = "Pass rate"; cells["B11"] = SafeGetDouble(latest, "pass_rate").ToString("P1");
                var fk = SafeGetStringArray(latest, "failure_keys");
                cells["A12"] = "Failure keys"; cells["B12"] = fk.Length > 0 ? string.Join(", ", fk) : "none";

                // Delta vs previous run
                if (runs.Count >= 2)
                {
                    var prev = runs[runs.Count - 2];
                    double prevRate = SafeGetDouble(prev, "pass_rate");
                    double currRate = SafeGetDouble(latest, "pass_rate");
                    cells["A14"] = "Previous pass rate"; cells["B14"] = prevRate.ToString("P1");
                    cells["A15"] = "Delta"; cells["B15"] = currRate > prevRate ? "IMPROVED" : currRate < prevRate ? "REGRESSED" : "unchanged";
                    var prevFk = new HashSet<string>(SafeGetStringArray(prev, "failure_keys"), StringComparer.OrdinalIgnoreCase);
                    var currFk = new HashSet<string>(SafeGetStringArray(latest, "failure_keys"), StringComparer.OrdinalIgnoreCase);
                    var fixedF = prevFk.Except(currFk).ToList();
                    var newF = currFk.Except(prevFk).ToList();
                    cells["A16"] = "Fixed"; cells["B16"] = fixedF.Count > 0 ? string.Join(", ", fixedF) : "none";
                    cells["A17"] = "New failures"; cells["B17"] = newF.Count > 0 ? string.Join(", ", newF) : "none";
                }

                sheets.Add(new { Name = "Overview", Rows = 100, Columns = 26, Cells = cells });
            }

            // --- Sheet 2: Run History ---
            {
                var cells = new Dictionary<string, string>();
                cells["A1"] = "#"; cells["B1"] = "Timestamp"; cells["C1"] = "Commit"; cells["D1"] = "Provider";
                cells["E1"] = "Model"; cells["F1"] = "Spec"; cells["G1"] = "Tests"; cells["H1"] = "Steps";
                cells["I1"] = "Pass rate"; cells["J1"] = "Failures";
                int startIdx = Math.Max(0, runs.Count - 50);
                int row = 2;
                for (int i = runs.Count - 1; i >= startIdx; i--)
                {
                    var r = runs[i];
                    string ts = SafeGetString(r, "timestamp") ?? "";
                    if (ts.Length > 19) ts = ts.Substring(0, 19);
                    var fk = SafeGetStringArray(r, "failure_keys");
                    cells[$"A{row}"] = (i + 1).ToString();
                    cells[$"B{row}"] = ts;
                    cells[$"C{row}"] = SafeGetString(r, "commit") ?? "";
                    cells[$"D{row}"] = SafeGetString(r, "provider") ?? "mock";
                    cells[$"E{row}"] = SafeGetString(r, "model") ?? "";
                    cells[$"F{row}"] = SafeGetString(r, "spec_file") ?? "";
                    cells[$"G{row}"] = $"{SafeGetInt(r, "tests_passed")}/{SafeGetInt(r, "tests_total")}";
                    cells[$"H{row}"] = $"{SafeGetInt(r, "steps_passed")}/{SafeGetInt(r, "steps_total")}";
                    cells[$"I{row}"] = SafeGetDouble(r, "pass_rate").ToString("P1");
                    cells[$"J{row}"] = fk.Length > 0 ? string.Join(", ", fk.Take(5)) : "";
                    row++;
                }
                sheets.Add(new { Name = "Run History", Rows = 100, Columns = 26, Cells = cells, FreezeTopRow = true });
            }

            // --- Sheet 3: Capabilities ---
            if (capMap.Count > 0)
            {
                var cells = new Dictionary<string, string>();
                cells["A1"] = "Capability"; cells["B1"] = "Tests"; cells["C1"] = "Covered"; cells["D1"] = "Passing"; cells["E1"] = "Failing"; cells["F1"] = "Status";
                int row = 2;
                foreach (var kv in capMap.OrderBy(x => x.Key))
                {
                    var ids = kv.Value.Where(id => !string.IsNullOrWhiteSpace(id)).ToList();
                    int covered = ids.Count;
                    int passing = 0, failing = 0;
                    foreach (var id in ids)
                    {
                        if (testResults.TryGetValue(id, out var p)) { if (p) passing++; else failing++; }
                    }
                    string status = covered == 0 ? "NO TESTS" : failing > 0 ? "FAIL" : passing == covered ? "PASS" : "PARTIAL";
                    cells[$"A{row}"] = kv.Key.Replace("_", " ");
                    cells[$"B{row}"] = string.Join(", ", ids);
                    cells[$"C{row}"] = covered.ToString();
                    cells[$"D{row}"] = passing.ToString();
                    cells[$"E{row}"] = failing.ToString();
                    cells[$"F{row}"] = status;
                    row++;
                }
                sheets.Add(new { Name = "Capabilities", Rows = 100, Columns = 26, Cells = cells, FreezeTopRow = true });
            }

            // --- Sheet 4: Failures ---
            {
                var cells = new Dictionary<string, string>();
                cells["A1"] = "Failure key"; cells["B1"] = "Occurrences"; cells["C1"] = "First seen"; cells["D1"] = "Last seen";
                if (allFailures.Count == 0)
                {
                    cells["A2"] = "(no failures recorded)";
                }
                else
                {
                    int row = 2;
                    foreach (var kv in allFailures.OrderByDescending(x => x.Value))
                    {
                        string? first = null, last = null;
                        foreach (var r in runs)
                        {
                            var fk = SafeGetStringArray(r, "failure_keys");
                            if (fk.Contains(kv.Key, StringComparer.OrdinalIgnoreCase))
                            {
                                string ts = SafeGetString(r, "timestamp") ?? "";
                                if (ts.Length > 10) ts = ts.Substring(0, 10);
                                if (first == null) first = ts;
                                last = ts;
                            }
                        }
                        cells[$"A{row}"] = kv.Key;
                        cells[$"B{row}"] = kv.Value.ToString();
                        cells[$"C{row}"] = first ?? "";
                        cells[$"D{row}"] = last ?? "";
                        row++;
                    }
                }
                sheets.Add(new { Name = "Failures", Rows = 100, Columns = 26, Cells = cells, FreezeTopRow = true });
            }

            // --- Sheet 5: Next Actions ---
            {
                var cells = new Dictionary<string, string>();
                cells["A1"] = "Priority"; cells["B1"] = "Action"; cells["C1"] = "Details";
                int row = 2;
                if (capMap.Count > 0)
                {
                    var uncovered = capMap.Where(kv => kv.Value.Count == 0).Select(kv => kv.Key.Replace("_", " ")).ToList();
                    foreach (var cap in uncovered)
                    {
                        cells[$"A{row}"] = "Add tests";
                        cells[$"B{row}"] = $"No tests for: {cap}";
                        cells[$"C{row}"] = "Create test workbooks covering this capability";
                        row++;
                    }
                    var failingCaps = capMap.Where(kv => kv.Value.Any(id => testResults.TryGetValue(id, out var p) && !p)).ToList();
                    foreach (var cap in failingCaps)
                    {
                        var failIds = cap.Value.Where(id => testResults.TryGetValue(id, out var p) && !p).ToList();
                        cells[$"A{row}"] = "Fix failing";
                        cells[$"B{row}"] = $"{cap.Key.Replace("_", " ")}";
                        cells[$"C{row}"] = $"Failing tests: {string.Join(", ", failIds)}";
                        row++;
                    }
                }
                if (allFailures.Count > 0)
                {
                    foreach (var kv in allFailures.OrderByDescending(x => x.Value).Take(5))
                    {
                        cells[$"A{row}"] = "Recurring failure";
                        cells[$"B{row}"] = kv.Key;
                        cells[$"C{row}"] = $"{kv.Value} occurrences across runs";
                        row++;
                    }
                }
                if (row == 2) { cells["A2"] = "—"; cells["B2"] = "All clear"; cells["C2"] = "No actions needed"; }
                sheets.Add(new { Name = "Next Actions", Rows = 100, Columns = 26, Cells = cells, FreezeTopRow = true });
            }

            var wb = new { FormatVersion = 1, Sheets = sheets };
            // Use default (PascalCase) naming to match SpreadsheetIO's WorkbookData loader
            string wbJson = JsonSerializer.Serialize(wb, new JsonSerializerOptions { WriteIndented = true });
            TryWriteAllText(Path.Combine(outputDir, "dashboard.workbook.json"), wbJson);
        }

        private static string? SafeGetString(Dictionary<string, JsonElement> d, string key)
        {
            if (d.TryGetValue(key, out var el))
            {
                if (el.ValueKind == JsonValueKind.String) return el.GetString();
                if (el.ValueKind == JsonValueKind.Null) return null;
                return el.GetRawText().Trim('"');
            }
            return null;
        }

        private static int SafeGetInt(Dictionary<string, JsonElement> d, string key)
        {
            if (d.TryGetValue(key, out var el) && el.ValueKind == JsonValueKind.Number) return el.GetInt32();
            return 0;
        }

        private static double SafeGetDouble(Dictionary<string, JsonElement> d, string key)
        {
            if (d.TryGetValue(key, out var el) && el.ValueKind == JsonValueKind.Number) return el.GetDouble();
            return 0.0;
        }

        private static string[] SafeGetStringArray(Dictionary<string, JsonElement> d, string key)
        {
            if (d.TryGetValue(key, out var el) && el.ValueKind == JsonValueKind.Array)
            {
                var list = new List<string>();
                foreach (var item in el.EnumerateArray())
                {
                    var s = item.GetString(); if (!string.IsNullOrWhiteSpace(s)) list.Add(s!);
                }
                return list.ToArray();
            }
            return Array.Empty<string>();
        }
    }
}
