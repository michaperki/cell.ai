using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace SpreadsheetApp.Core.AI
{
    // Simple heuristic provider for development without network calls.
    public sealed class MockInferenceProvider : IInferenceProvider
    {
        public Task<AIResult> GenerateFillAsync(AIContext context, CancellationToken cancellationToken)
        {
            // Trim prompt for matching
            var prompt = (context.Prompt ?? string.Empty).Trim();
            var title = (context.Title ?? string.Empty).Trim();
            var text = (title + "\n" + prompt).ToLowerInvariant();

            var rows = Math.Max(1, context.Rows);
            var cols = Math.Max(1, context.Cols);

            var cells = new string[rows][];
            for (int r = 0; r < rows; r++) cells[r] = new string[cols];

            List<string> items = new();

            if (text.Contains("caesar") && text.Contains("salad"))
            {
                items = new List<string>
                {
                    "Romaine lettuce",
                    "Parmesan cheese",
                    "Croutons",
                    "Caesar dressing",
                    "Chicken (optional)",
                    "Lemon juice",
                    "Olive oil",
                    "Garlic",
                    "Anchovies",
                    "Egg yolk",
                    "Dijon mustard"
                };
            }
            else if (text.Contains("grocery") || text.Contains("shopping"))
            {
                items = new List<string>
                {
                    "Milk",
                    "Eggs",
                    "Bread",
                    "Butter",
                    "Chicken breast",
                    "Rice",
                    "Pasta",
                    "Tomatoes",
                    "Bananas",
                    "Apples",
                    "Onions"
                };
            }
            else if (text.Contains("todo") || text.Contains("to-do"))
            {
                items = new List<string> { "Plan", "Draft", "Review", "Revise", "Ship" };
            }
            else
            {
                // Generic items
                int total = rows * cols;
                for (int i = 1; i <= total; i++) items.Add($"Item {i}");
            }

            // Fill by rows then cols
            int k = 0;
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    if (k < items.Count) cells[r][c] = items[k++];
                    else cells[r][c] = string.Empty;
                }
            }

            return Task.FromResult(new AIResult { Cells = cells });
        }
    }
}

