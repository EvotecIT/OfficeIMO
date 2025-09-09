using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace OfficeIMO.Excel
{
    /// <summary>
    /// Utility helpers for parsing and converting Excel A1 references. Public so examples and
    /// consumers can reuse consistent logic without re-implementing regexes or math.
    /// </summary>
    public static class A1
    {
        private static readonly Regex RangeRx = new("^\\s*([A-Za-z]+)(\\d+)\\s*:\\s*([A-Za-z]+)(\\d+)\\s*$", RegexOptions.Compiled);
        private static readonly Regex CellRx = new("^\\s*([A-Za-z]+)(\\d+)\\s*$", RegexOptions.Compiled);

        /// <summary>
        /// Parses a single A1 cell reference (e.g., "B5") into a 1-based (row, column) tuple.
        /// Returns (0,0) when the input does not match a valid simple cell reference.
        /// </summary>
        /// <param name="cellRef">A1 cell reference, without sheet prefix.</param>
        /// <returns>Tuple of row and column (1-based). Returns (0,0) if invalid.</returns>
        public static (int Row, int Col) ParseCellRef(string cellRef)
        {
            if (string.IsNullOrWhiteSpace(cellRef)) return (0, 0);
            var m = CellRx.Match(cellRef);
            if (!m.Success) return (0, 0);
            var col = ColumnLettersToIndex(m.Groups[1].Value);
            var row = int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture);
            return (row, col);
        }

        /// <summary>
        /// Parses an A1 range (e.g., "A1:B10") into 1-based, normalized bounds.
        /// If the bounds are inverted, they are swapped so that r1 &lt;= r2 and c1 &lt;= c2.
        /// </summary>
        /// <param name="a1Range">A1 range string, without sheet prefix.</param>
        /// <returns>(r1, c1, r2, c2) 1-based coordinates.</returns>
        /// <exception cref="ArgumentException">Thrown when the input is not a valid A1 range.</exception>
        /// <example>
        /// var (r1, c1, r2, c2) = A1.ParseRange("B2:D10");
        /// // r1=2, c1=2, r2=10, c2=4
        /// </example>
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

        /// <summary>
        /// Converts column letters (e.g., "A", "AA") to a 1-based column index.
        /// Non-letter characters are ignored; returns 0 for empty/invalid input.
        /// </summary>
        /// <param name="letters">Excel column letters.</param>
        /// <returns>1-based column index, or 0 when input yields no letters.</returns>
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

        /// <summary>
        /// Converts a 1-based column index to Excel column letters (e.g., 1→"A", 27→"AA").
        /// </summary>
        /// <param name="index">1-based column index.</param>
        /// <returns>Excel column letters; returns "A" for non-positive inputs.</returns>
        /// <example>
        /// string col = A1.ColumnIndexToLetters(28); // "AB"
        /// int idx = A1.ColumnLettersToIndex("AB"); // 28
        /// </example>
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
