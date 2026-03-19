using System;
using System.IO;
using System.Text.Json;
using System.Security.Cryptography;
using System.Text;

namespace SpreadsheetApp.Core
{
    public sealed class AppSettings
    {
        public bool AiEnabled { get; set; } = true;
        public string Provider { get; set; } = "Auto"; // Auto, Mock, OpenAI, Anthropic, External
        public string? ExternalApiBaseUrl { get; set; } = null; // Full endpoint URL for GenerateFill
        public string? ApiKeyProtectedBase64 { get; set; } = null; // DPAPI-protected key

        public static string SettingsPath()
        {
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SpreadsheetApp");
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, "settings.json");
        }

        public static AppSettings Load()
        {
            try
            {
                var path = SettingsPath();
                if (File.Exists(path))
                {
                    var json = File.ReadAllText(path);
                    var s = JsonSerializer.Deserialize<AppSettings>(json);
                    if (s != null) return s;
                }
            }
            catch { }
            return new AppSettings();
        }

        public void Save()
        {
            try
            {
                var json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(SettingsPath(), json);
            }
            catch { }
        }

        // --- API key helpers (Windows DPAPI) ---
        private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("SpreadsheetApp-AI-Entropy-v1");

        public bool HasApiKey => !string.IsNullOrEmpty(ApiKeyProtectedBase64);

        public void SetApiKey(string plainText)
        {
            if (string.IsNullOrEmpty(plainText)) { ApiKeyProtectedBase64 = null; return; }
            try
            {
                var bytes = Encoding.UTF8.GetBytes(plainText);
                var protectedBytes = ProtectedData.Protect(bytes, Entropy, DataProtectionScope.CurrentUser);
                ApiKeyProtectedBase64 = Convert.ToBase64String(protectedBytes);
            }
            catch { ApiKeyProtectedBase64 = null; }
        }

        public string? GetApiKey()
        {
            try
            {
                if (string.IsNullOrEmpty(ApiKeyProtectedBase64)) return null;
                var protectedBytes = Convert.FromBase64String(ApiKeyProtectedBase64);
                var bytes = ProtectedData.Unprotect(protectedBytes, Entropy, DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(bytes);
            }
            catch { return null; }
        }

        public void ClearApiKey()
        {
            ApiKeyProtectedBase64 = null;
        }
    }
}
