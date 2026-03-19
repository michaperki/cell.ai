using System;
using System.Windows.Forms;

namespace SpreadsheetApp
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            // Load .env first so provider selection can use environment variables
            SpreadsheetApp.Core.Env.LoadDotEnv();
            ApplicationConfiguration.Initialize();
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += (s, e) =>
            {
                try { MessageBox.Show($"Unhandled UI exception:\n{e.Exception}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                catch { }
            };
            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            {
                try { MessageBox.Show($"Unhandled exception:\n{e.ExceptionObject}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                catch { }
            };
            try
            {
                Application.Run(new UI.MainForm());
            }
            catch (Exception ex)
            {
                try { MessageBox.Show($"Fatal error:\n{ex}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                catch { }
            }
        }
    }
}
