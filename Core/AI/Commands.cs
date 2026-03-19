using System.Collections.Generic;

namespace SpreadsheetApp.Core.AI
{
    public enum AICommandType { SetValues, SetTitle, CreateSheet }

    public interface IAICommand
    {
        AICommandType Type { get; }
        string Summarize();
    }

    public sealed class SetValuesCommand : IAICommand
    {
        public AICommandType Type => AICommandType.SetValues;
        public int StartRow { get; set; }
        public int StartCol { get; set; }
        public string[][] Values { get; set; } = new string[0][];
        public string Summarize() => $"Set {Values.Length}x{(Values.Length>0?Values[0].Length:0)} values at {StartRow+1},{StartCol+1}";
    }

    public sealed class SetTitleCommand : IAICommand
    {
        public AICommandType Type => AICommandType.SetTitle;
        public int StartRow { get; set; }
        public int StartCol { get; set; }
        public int Rows { get; set; } = 1;
        public int Cols { get; set; } = 1;
        public string Text { get; set; } = string.Empty;
        public string Summarize() => $"Title '{Text}' at {StartRow+1},{StartCol+1} ({Rows}x{Cols})";
    }

    public sealed class AIPlan
    {
        public List<IAICommand> Commands { get; } = new();
    }

    public sealed class CreateSheetCommand : IAICommand
    {
        public AICommandType Type => AICommandType.CreateSheet;
        public string Name { get; set; } = "New Sheet";
        public string Summarize() => $"Create sheet '{Name}'";
    }
}
