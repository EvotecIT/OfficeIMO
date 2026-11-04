using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Excel {
    internal static class ExcelColor {
        /// <summary>Creates a native SpreadsheetML color, retaining supported theme slots. Callers must resolve status roles against their theme first.</summary>
        internal static DocumentFormat.OpenXml.Spreadsheet.Color ToSpreadsheetColor(Color color) {
            uint? themeIndex = color.ThemeSlot switch {
                "lt1" => 0U,
                "dk1" => 1U,
                "lt2" => 2U,
                "dk2" => 3U,
                "accent1" => 4U,
                "accent2" => 5U,
                "accent3" => 6U,
                "accent4" => 7U,
                "accent5" => 8U,
                "accent6" => 9U,
                _ => null
            };
            return themeIndex.HasValue
                ? new DocumentFormat.OpenXml.Spreadsheet.Color { Theme = themeIndex.Value }
                : new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = ToArgbHex(color) };
        }

        /// <summary>
        /// Converts OfficeIMO color to ARGB hex string expected by Spreadsheet colors (AARRGGBB, uppercase).
        /// </summary>
        public static string ToArgbHex(Color color) {
            var hex = color.ToHex().ToUpperInvariant(); // RRGGBBAA or RRGGBB
            if (hex.Length >= 8) {
                string rrggbb = hex.Substring(0, 6);
                string aa = hex.Substring(6, 2);
                return aa + rrggbb;
            }
            if (hex.Length == 6) {
                return "FF" + hex;
            }
            return "FFFFFFFF";
        }
    }
}

