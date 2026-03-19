using System;

namespace SpreadsheetApp.Core
{
    public static class CellAddress
    {
        public static string ToAddress(int row, int col)
        {
            return ColumnIndexToName(col) + (row + 1).ToString();
        }

        public static string ColumnIndexToName(int col)
        {
            // 0 -> A, 25 -> Z, 26 -> AA
            col += 1;
            string name = string.Empty;
            while (col > 0)
            {
                int rem = (col - 1) % 26;
                name = (char)('A' + rem) + name;
                col = (col - 1) / 26;
            }
            return name;
        }

        public static bool TryParse(string addr, out int row, out int col)
        {
            row = -1; col = -1;
            if (string.IsNullOrWhiteSpace(addr)) return false;

            int i = 0;
            // Letters
            int lettersStart = 0;
            while (i < addr.Length && char.IsLetter(addr[i])) i++;
            if (i == 0) return false;
            string letters = addr.Substring(lettersStart, i - lettersStart).ToUpperInvariant();
            // Digits
            int digitsStart = i;
            while (i < addr.Length && char.IsDigit(addr[i])) i++;
            if (digitsStart == i) return false;
            if (i != addr.Length) return false; // extra chars

            col = ColumnNameToIndex(letters);
            if (!int.TryParse(addr.Substring(digitsStart), out int row1)) return false;
            if (row1 <= 0) return false;
            row = row1 - 1;
            return true;
        }

        public static int ColumnNameToIndex(string name)
        {
            int col = 0;
            foreach (char ch in name.ToUpperInvariant())
            {
                if (ch < 'A' || ch > 'Z') throw new ArgumentException("Invalid column name");
                col = col * 26 + (ch - 'A' + 1);
            }
            return col - 1;
        }
    }
}

