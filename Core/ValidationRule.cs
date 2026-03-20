namespace SpreadsheetApp.Core
{
    public enum ValidationMode { None, NumberBetween, List }

    public sealed class ValidationRule
    {
        public ValidationMode Mode { get; set; } = ValidationMode.None;
        public bool AllowEmpty { get; set; } = true;
        public double? Min { get; set; }
        public double? Max { get; set; }
        public string[]? AllowedList { get; set; }

        public bool IsEmpty()
        {
            if (Mode == ValidationMode.None) return true;
            if (Mode == ValidationMode.List && (AllowedList == null || AllowedList.Length == 0)) return true;
            return false;
        }
    }
}

