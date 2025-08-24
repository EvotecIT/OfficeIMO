using SixLabors.ImageSharp;

namespace OfficeIMO.Excel
{
    internal static class ExcelColor
    {
        /// <summary>
        /// Converts SixLabors Color to ARGB hex string expected by Spreadsheet colors (AARRGGBB, uppercase).
        /// </summary>
        public static string ToArgbHex(Color color)
        {
            var hex = color.ToHex().ToUpperInvariant(); // RRGGBBAA or RRGGBB
            if (hex.Length >= 8)
            {
                string rrggbb = hex.Substring(0, 6);
                string aa = hex.Substring(6, 2);
                return aa + rrggbb;
            }
            if (hex.Length == 6)
            {
                return "FF" + hex;
            }
            return "FFFFFFFF";
        }
    }
}

