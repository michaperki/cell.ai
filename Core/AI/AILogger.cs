using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace SpreadsheetApp.Core.AI
{
    public static class AILogger
    {
        private static bool IsEnabled()
        {
            var v = Environment.GetEnvironmentVariable("AI_DEBUG_LOG");
            if (string.IsNullOrWhiteSpace(v)) return false;
            v = v.Trim().ToLowerInvariant();
            return v == "1" || v == "true" || v == "yes";
        }

        private static string GetLogDir()
        {
            var dir = Environment.GetEnvironmentVariable("AI_DEBUG_DIR");
            if (!string.IsNullOrWhiteSpace(dir)) return dir;
            try
            {
                string baseDir = AppDomain.CurrentDomain.BaseDirectory ?? Environment.CurrentDirectory;
                return Path.Combine(baseDir, "logs", "ai");
            }
            catch
            {
                return Path.Combine(Environment.CurrentDirectory, "logs", "ai");
            }
        }

        public static void Log(string surface, string? provider, string? model, AIUsage? usage, int? latencyMs,
                               string? sheet, int? startRow, int? startCol, int? rows, int? cols,
                               string? userPrompt, string? systemPrompt,
                               int? commandCount, int? writeCells)
        {
            try
            {
                if (!IsEnabled()) return;
                bool includePrompts = false;
                var vp = Environment.GetEnvironmentVariable("AI_DEBUG_PROMPT");
                if (!string.IsNullOrWhiteSpace(vp))
                {
                    vp = vp.Trim().ToLowerInvariant();
                    includePrompts = vp == "1" || vp == "true" || vp == "yes";
                }
                string dir = GetLogDir();
                Directory.CreateDirectory(dir);
                var obj = new Dictionary<string, object?>
                {
                    ["ts"] = DateTime.UtcNow.ToString("o"),
                    ["surface"] = surface,
                    ["provider"] = provider,
                    ["model"] = model,
                    ["latency_ms"] = latencyMs,
                    ["sheet"] = sheet,
                    ["start_row"] = startRow,
                    ["start_col"] = startCol,
                    ["rows"] = rows,
                    ["cols"] = cols,
                    ["input_tokens"] = usage?.InputTokens,
                    ["output_tokens"] = usage?.OutputTokens,
                    ["total_tokens"] = usage?.TotalTokens,
                    ["context_limit"] = usage?.ContextLimit,
                    ["remaining_context"] = usage?.RemainingContext,
                    ["command_count"] = commandCount,
                    ["write_cells"] = writeCells,
                    ["user_prompt_len"] = userPrompt?.Length,
                    ["system_prompt_len"] = systemPrompt?.Length
                };
                if (includePrompts)
                {
                    obj["user_prompt"] = userPrompt;
                    obj["system_prompt"] = systemPrompt;
                }
                string json = JsonSerializer.Serialize(obj);
                string path = Path.Combine(dir, DateTime.UtcNow.ToString("yyyyMMdd") + ".jsonl");
                File.AppendAllText(path, json + Environment.NewLine);
            }
            catch { }
        }
    }
}

