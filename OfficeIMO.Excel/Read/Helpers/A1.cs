using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace OfficeIMO.Excel.Read
{
    internal static class A1
    {
        private static readonly Regex RangeRx = new("^\\s*([A-Za-z]+)(\\d+)\\s*:\\s*([A-Za-z]+)(\\d+)\\s*$", RegexOptions.Compiled);
        private static readonly Regex CellRx = new("^\\s*([A-Za-z]+)(\\d+)\\s*$", RegexOptions.Compiled);

        public static (int Row, int Col) ParseCellRef(string cellRef)
        {
            if (string.IsNullOrWhiteSpace(cellRef)) return (0, 0);
            var m = CellRx.Match(cellRef);
            if (!m.Success) return (0, 0);
            var col = ColumnLettersToIndex(m.Groups[1].Value);
            var row = int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture);
            return (row, col);
        }

        public static (int r1, int c1, int r2, int c2) ParseRange(string a1Range)
        {
            var m = RangeRx.Match(a1Range);
            if (!m.Success) throw new ArgumentException($"Invalid A1 range '{a1Range}'.");
            var c1 = ColumnLettersToIndex(m.Groups[1].Value);
            var r1 = int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture);
            var c2 = ColumnLettersToIndex(m.Groups[3].Value);
            var r2 = int.Parse(m.Groups[4].Value, CultureInfo.InvariantCulture);
            if (c1 > c2) (c1, c2) = (c2, c1);
            if (r1 > r2) (r1, r2) = (r2, r1);
            return (r1, c1, r2, c2);
        }

        public static int ColumnLettersToIndex(string letters)
        {
            int res = 0;
            foreach (char ch in letters.ToUpperInvariant())
            {
                if (ch < 'A' || ch > 'Z') continue;
                res = res * 26 + (ch - 'A' + 1);
            }
            return res;
        }

        public static string ColumnIndexToLetters(int index)
        {
            if (index <= 0) return "A";
            string letters = string.Empty;
            int n = index;
            while (n > 0)
            {
                int rem = (n - 1) % 26;
                letters = (char)('A' + rem) + letters;
                n = (n - 1) / 26;
            }
            return letters;
        }
    }
}
