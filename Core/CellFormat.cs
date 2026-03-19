namespace SpreadsheetApp.Core
{
    public enum CellHAlign { Left, Center, Right }

    public class CellFormat
    {
        public bool Bold { get; set; }
        public int? ForeColorArgb { get; set; }
        public int? BackColorArgb { get; set; }
        public string? NumberFormat { get; set; } // e.g., "General", "0.00"
        public CellHAlign HAlign { get; set; } = CellHAlign.Left;

        public bool IsDefault()
        {
            return !Bold && ForeColorArgb == null && BackColorArgb == null && string.IsNullOrEmpty(NumberFormat) && HAlign == CellHAlign.Left;
        }
    }
}

