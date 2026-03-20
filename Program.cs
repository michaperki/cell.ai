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
