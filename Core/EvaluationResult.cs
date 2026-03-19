using System.Globalization;

namespace SpreadsheetApp.Core
{
    public class EvaluationResult
    {
        public double? Number { get; }
        public string? Text { get; }
        public string? Error { get; }

        private EvaluationResult(double? number, string? text, string? error)
        {
            Number = number; Text = text; Error = error;
        }

        public static EvaluationResult FromNumber(double d) => new EvaluationResult(d, null, null);
        public static EvaluationResult FromText(string? s) => new EvaluationResult(null, s ?? string.Empty, null);
        public static EvaluationResult FromError(string e) => new EvaluationResult(null, null, e);

        public string ToDisplay()
        {
            if (Error != null) return "#ERR: " + Error;
            if (Number is double d) return d.ToString("G", CultureInfo.InvariantCulture);
            return Text ?? string.Empty;
        }
    }
}

