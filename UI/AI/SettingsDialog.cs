using System;
using System.Windows.Forms;
using SpreadsheetApp.Core;
using SpreadsheetApp.UI;

namespace SpreadsheetApp.UI.AI
{
    public sealed class SettingsDialog : Form
    {
        private readonly CheckBox _chkEnableAi = new() { Text = "Enable AI features", AutoSize = true };
        private readonly ComboBox _cmbProvider = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 220 };
        private readonly Label _lblKeyHint = new() { AutoSize = true, ForeColor = Theme.TextMuted };
        private readonly Button _btnTest = new() { Text = "Test Connection" };
        private readonly Button _btnOK = new() { Text = "OK", DialogResult = DialogResult.OK };
        private readonly Button _btnCancel = new() { Text = "Cancel", DialogResult = DialogResult.Cancel };

        public bool AiEnabled { get; set; }
        public string Provider { get; set; } = "Mock";

        public SettingsDialog(AppSettings settings)
        {
            Text = "AI Settings";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            ClientSize = new System.Drawing.Size(420, 220);
            MinimizeBox = false; MaximizeBox = false;
            BackColor = System.Drawing.Color.White;
            Font = Theme.UI;

            Theme.StyleSecondary(_btnTest);
            Theme.StylePrimary(_btnOK);
            Theme.StyleGhost(_btnCancel);

            _chkEnableAi.ForeColor = Theme.TextPrimary;

            var lblProvider = new Label { Text = "Provider:", AutoSize = true, Left = 20, Top = 60, ForeColor = Theme.TextPrimary };
            _chkEnableAi.Left = 20; _chkEnableAi.Top = 20;
            _cmbProvider.Left = 120; _cmbProvider.Top = 56;
            _cmbProvider.Items.AddRange(new object[] { "Auto (env)", "OpenAI", "Anthropic", "Mock (local)", "External API" });
            _lblKeyHint.Left = 20; _lblKeyHint.Top = 92; _lblKeyHint.Text = "API keys are loaded from .env or system environment variables.";

            _btnTest.Left = 120; _btnTest.Top = 185;
            _btnOK.Left = 240; _btnOK.Top = 185;
            _btnCancel.Left = 320; _btnCancel.Top = 185;

            Controls.AddRange(new Control[] { _chkEnableAi, lblProvider, _cmbProvider, _lblKeyHint, _btnTest, _btnOK, _btnCancel });
            AcceptButton = _btnOK; CancelButton = _btnCancel;

            // Initialize from settings
            AiEnabled = settings.AiEnabled;
            Provider = settings.Provider;
            _chkEnableAi.Checked = AiEnabled;
            _cmbProvider.SelectedIndex = Provider switch { "OpenAI" => 1, "Anthropic" => 2, "Mock" => 3, "External" => 4, _ => 0 };

            _btnOK.Click += (_, __) =>
            {
                AiEnabled = _chkEnableAi.Checked;
                Provider = _cmbProvider.SelectedIndex switch { 0 => "Auto", 1 => "OpenAI", 2 => "Anthropic", 3 => "Mock", 4 => "External", _ => "Auto" };
                DialogResult = DialogResult.OK;
            };

            _btnTest.Click += async (_, __) =>
            {
                try
                {
                    // Reload .env in case the file was added after app start
                    SpreadsheetApp.Core.Env.LoadDotEnv();
                    // Test by environment keys regardless of dropdown selection
                    var openAiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY");
                    var anthropicKey = Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY");
                    string resolved;
                    SpreadsheetApp.Core.AI.IInferenceProvider provider;
                    if (!string.IsNullOrWhiteSpace(openAiKey))
                    {
                        resolved = "OpenAI";
                        provider = new SpreadsheetApp.Core.AI.OpenAIProvider(openAiKey);
                    }
                    else if (!string.IsNullOrWhiteSpace(anthropicKey))
                    {
                        resolved = "Anthropic";
                        provider = new SpreadsheetApp.Core.AI.AnthropicProvider(anthropicKey);
                    }
                    else
                    {
                        resolved = "Mock";
                        provider = new SpreadsheetApp.Core.AI.MockInferenceProvider();
                        var baseDir = AppDomain.CurrentDomain.BaseDirectory ?? Environment.CurrentDirectory;
                        var envPath = System.IO.Path.Combine(baseDir, ".env");
                        MessageBox.Show(this, $"Using Mock provider (no network).\nNo OPENAI_API_KEY or ANTHROPIC_API_KEY detected.\nAdd them to system/user env or create:\n{envPath}\nThen restart the app.", "Test Connection", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    using var cts = new System.Threading.CancellationTokenSource(TimeSpan.FromSeconds(10));
                    var t0 = DateTime.UtcNow;
                    var res = await provider.GenerateFillAsync(new SpreadsheetApp.Core.AI.AIContext { SheetName = "Test", StartRow = 0, StartCol = 0, Rows = 1, Cols = 1, Title = "Test", Prompt = "ok" }, cts.Token);
                    var ms = (int)(DateTime.UtcNow - t0).TotalMilliseconds;
                    var cell = (res.Cells.Length > 0 && res.Cells[0].Length > 0) ? res.Cells[0][0] : "";
                    MessageBox.Show(this, $"{resolved} OK ({ms} ms). Sample: '{cell}'.", "Test Connection", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (OperationCanceledException)
                {
                    MessageBox.Show(this, "Timed out.", "Test Connection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, $"Failed: {ex.Message}", "Test Connection", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
        }
    }
}
