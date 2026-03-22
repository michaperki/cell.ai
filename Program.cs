using System;
using System.Windows.Forms;
using System.Linq;
using SpreadsheetApp.Core;

namespace SpreadsheetApp
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            // Load .env first so provider selection can use environment variables
            SpreadsheetApp.Core.Env.LoadDotEnv();
            // Lightweight CLI: export docs index and exit
            try
            {
                var args = Environment.GetCommandLineArgs();
                // Headless test runner: run full suite or a single test without launching UI
                // Usage:
                //   --run-tests [tests/TEST_SPECS.json] [--output-dir tests/output]
                //   --run-one <tests/test_XX_....workbook.json> [--output-dir tests/output]
                if (args.Any(a => string.Equals(a, "--run-tests", StringComparison.OrdinalIgnoreCase)))
                {
                    string specs = "tests/TEST_SPECS.json";
                    string outputDir = System.IO.Path.Combine("tests", "output");
                    bool reflection = false;
                    for (int i = 0; i < args.Length; i++)
                    {
                        if (string.Equals(args[i], "--run-tests", StringComparison.OrdinalIgnoreCase))
                        {
                            if (i + 1 < args.Length && !args[i + 1].StartsWith("--", StringComparison.Ordinal))
                                specs = args[i + 1];
                        }
                        else if (string.Equals(args[i], "--output-dir", StringComparison.OrdinalIgnoreCase))
                        {
                            if (i + 1 < args.Length) outputDir = args[i + 1];
                        }
                        else if (string.Equals(args[i], "--reflection", StringComparison.OrdinalIgnoreCase))
                        {
                            reflection = true;
                        }
                    }
                    int code = SpreadsheetApp.Core.HeadlessTestRunner.RunAll(specs, outputDir, reflection);
                    try { Console.WriteLine($"Headless test run completed with exit code {code}."); } catch { }
                    Environment.ExitCode = code; return;
                }
                if (args.Any(a => string.Equals(a, "--bless-baseline", StringComparison.OrdinalIgnoreCase)))
                {
                    string outputDir = System.IO.Path.Combine("tests", "output");
                    for (int i = 0; i < args.Length; i++)
                    {
                        if (string.Equals(args[i], "--output-dir", StringComparison.OrdinalIgnoreCase))
                        {
                            if (i + 1 < args.Length) outputDir = args[i + 1];
                        }
                    }
                    try
                    {
                        string src = System.IO.Path.Combine(outputDir, "results_summary.json");
                        string dst = System.IO.Path.Combine(outputDir, "baseline.json");
                        if (!System.IO.File.Exists(src))
                        {
                            try { Console.Error.WriteLine($"results_summary.json not found in {outputDir}. Run tests first, then bless."); } catch { }
                            Environment.ExitCode = 2; return;
                        }
                        System.IO.Directory.CreateDirectory(outputDir);
                        System.IO.File.Copy(src, dst, overwrite: true);
                        try { Console.WriteLine($"Blessed baseline: {dst}"); } catch { }
                        Environment.ExitCode = 0; return;
                    }
                    catch (Exception ex)
                    {
                        try { Console.Error.WriteLine($"Failed to bless baseline: {ex.Message}"); } catch { }
                        Environment.ExitCode = 1; return;
                    }
                }
                if (args.Any(a => string.Equals(a, "--run-one", StringComparison.OrdinalIgnoreCase)))
                {
                    string file = string.Empty;
                    string outputDir = System.IO.Path.Combine("tests", "output");
                    bool reflection = false;
                    for (int i = 0; i < args.Length; i++)
                    {
                        if (string.Equals(args[i], "--run-one", StringComparison.OrdinalIgnoreCase))
                        {
                            if (i + 1 < args.Length && !args[i + 1].StartsWith("--", StringComparison.Ordinal))
                                file = args[i + 1];
                        }
                        else if (string.Equals(args[i], "--output-dir", StringComparison.OrdinalIgnoreCase))
                        {
                            if (i + 1 < args.Length) outputDir = args[i + 1];
                        }
                        else if (string.Equals(args[i], "--reflection", StringComparison.OrdinalIgnoreCase))
                        {
                            reflection = true;
                        }
                    }
                    if (string.IsNullOrWhiteSpace(file))
                    {
                        try { Console.Error.WriteLine("--run-one requires a .workbook.json file path"); } catch { }
                        Environment.ExitCode = 2; return;
                    }
                    int code = SpreadsheetApp.Core.HeadlessTestRunner.RunSingle(file, outputDir, reflection);
                    try { Console.WriteLine($"Headless single test completed with exit code {code}."); } catch { }
                    Environment.ExitCode = code; return;
                }
                if (args.Any(a => string.Equals(a, "--export-docs", StringComparison.OrdinalIgnoreCase)))
                {
                    var idx = DocsIndexer.Build(null);
                    string path = DocsIndexer.WriteJson(idx, null);
                    try { Console.WriteLine($"Exported docs JSON to: {path}"); } catch { }
                    return;
                }
            }
            catch { }
            ApplicationConfiguration.Initialize();
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += (s, e) =>
            {
                try { System.IO.File.AppendAllText("crash.log", $"[{DateTime.Now:o}] UI thread exception:\n{e.Exception}\n\n"); } catch { }
                try { MessageBox.Show($"Unhandled UI exception:\n{e.Exception}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                catch { }
            };
            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            {
                try { System.IO.File.AppendAllText("crash.log", $"[{DateTime.Now:o}] AppDomain exception:\n{e.ExceptionObject}\n\n"); } catch { }
                try { MessageBox.Show($"Unhandled exception:\n{e.ExceptionObject}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                catch { }
            };
            try
            {
                Application.Run(new UI.MainForm());
            }
            catch (Exception ex)
            {
                try { System.IO.File.AppendAllText("crash.log", $"[{DateTime.Now:o}] Fatal:\n{ex}\n\n"); } catch { }
                try { MessageBox.Show($"Fatal error:\n{ex}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                catch { }
            }
        }
    }
}
