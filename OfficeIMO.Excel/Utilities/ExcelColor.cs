using Color = OfficeIMO.OfficeColor;
using SpreadsheetColor = DocumentFormat.OpenXml.Spreadsheet.Color;

namespace OfficeIMO.Excel {
    internal static class ExcelColor {
        /// <summary>Creates a native SpreadsheetML color, retaining supported theme slots and resolving status roles.</summary>
        internal static SpreadsheetColor ToSpreadsheetColor(Color color, OfficeTheme? theme = null) {
            uint? index = OfficeOpenXmlThemeColorResolver.GetSpreadsheetThemeColorIndex(color.ThemeSlot);
            return index.HasValue
                ? new SpreadsheetColor { Theme = index.Value }
                : new SpreadsheetColor { Rgb = color.Resolve(theme ?? OfficeTheme.Default).ToArgbHex() };
        }

        /// <summary>Projects rich-run colors, preserving unmodified native roles and interpreting native literals as ARGB.</summary>
        internal static Color? FromSpreadsheetColor(SpreadsheetColor? color, string? resolvedArgb) {
            if (color?.Theme?.Value is uint index && (color.Tint?.Value ?? 0D) == 0D) {
                Color? role = index switch {
                    0U => Color.Background, 1U => Color.Text, 3U => Color.TextMuted,
                    4U => Color.Accent1, 5U => Color.Accent2, 6U => Color.Accent3,
                    7U => Color.Accent4, 8U => Color.Accent5, 9U => Color.Accent6, _ => null
                };
                if (role.HasValue) return role;
            }
            string? argb = resolvedArgb ?? color?.Rgb?.Value;
            if (string.IsNullOrWhiteSpace(argb)) return null;
            argb = argb!.TrimStart('#');
            return Color.ParseHex(argb.Length == 8 ? argb.Substring(2) + argb.Substring(0, 2) : argb);
        }

        /// <summary>Converts a literal color to native AARRGGBB, distinct from the Core RRGGBBAA projection.</summary>
        public static string ToArgbHex(Color color) => color.ToArgbHex();
    }
}
