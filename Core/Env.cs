using System;
using System.Collections.Generic;
using System.IO;

namespace SpreadsheetApp.Core
{
    public static class Env
    {
        public static void LoadDotEnv()
        {
            try
            {
                string baseDir = AppDomain.CurrentDomain.BaseDirectory ?? Environment.CurrentDirectory;
                string path = Path.Combine(baseDir, ".env");
                if (!File.Exists(path)) return;
                foreach (var line in File.ReadAllLines(path))
                {
                    var trimmed = line.Trim();
                    if (string.IsNullOrEmpty(trimmed)) continue;
                    if (trimmed.StartsWith("#")) continue;
                    int eq = trimmed.IndexOf('=');
                    if (eq <= 0) continue;
                    string key = trimmed.Substring(0, eq).Trim();
                    string value = trimmed.Substring(eq + 1).Trim();
                    if (value.Length >= 2 && ((value.StartsWith("\"") && value.EndsWith("\"")) || (value.StartsWith("'") && value.EndsWith("'"))))
                    {
                        value = value.Substring(1, value.Length - 2);
                    }
                    if (string.IsNullOrEmpty(key)) continue;
                    Environment.SetEnvironmentVariable(key, value);
                }
            }
            catch { }
        }
    }
}

