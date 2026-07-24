using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Excel.Utilities {
    /// <summary>Resolves SpreadsheetML and DrawingML colors through the shared Office color contracts.</summary>
    internal static class ExcelThemeColorResolver {
        private static readonly string?[] IndexedColors = {
            "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
            "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
            "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
            "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
            "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
            "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
            "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
            "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333"
        };

        internal static string? Resolve(OpenXmlElement? colorElement, WorkbookPart? workbookPart) =>
            colorElement is ColorType color ? Resolve(color, workbookPart) : null;

        internal static string? Resolve(
            ColorType? color,
            WorkbookPart? workbookPart,
            IReadOnlyDictionary<uint, string>? indexedColorOverrides = null) {
            if (color == null) {
                return null;
            }

            string? argb = NormalizeArgb(color.Rgb?.Value);
            if (argb == null && color.Theme?.Value is uint themeIndex) {
                OfficeColor? themed = OfficeOpenXmlThemeColorResolver.ResolveSpreadsheetThemeColor(
                    GetColorScheme(workbookPart),
                    themeIndex);
                argb = themed?.ToArgbHex();
            }

            if (argb == null && color.Indexed?.Value is uint indexed) {
                if (indexedColorOverrides != null && indexedColorOverrides.TryGetValue(indexed, out string? custom)) {
                    argb = NormalizeArgb(custom);
                } else if (indexed < IndexedColors.Length) {
                    argb = NormalizeArgb(IndexedColors[indexed]);
                }
            }

            if (argb == null) {
                return null;
            }

            return color.Tint?.Value is double tint && Math.Abs(tint) > double.Epsilon
                ? ApplySpreadsheetTint(argb, tint)
                : argb;
        }

        internal static string? Resolve(A.SolidFill? solidFill, WorkbookPart? workbookPart) {
            OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveColor(solidFill, GetColorScheme(workbookPart));
            return color?.ToArgbHex();
        }

        internal static string? ResolveTheme(uint themeIndex, double? tint, WorkbookPart? workbookPart) {
            OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveSpreadsheetThemeColor(
                GetColorScheme(workbookPart),
                themeIndex);
            if (!color.HasValue) {
                return null;
            }

            return tint.HasValue && Math.Abs(tint.Value) > double.Epsilon
                ? OfficeColorTransforms.SpreadsheetTint(color.Value, NormalizeSpreadsheetTint(tint.Value)).ToArgbHex()
                : color.Value.ToArgbHex();
        }

        internal static string ApplySpreadsheetTint(string argb, double tint) {
            if (!TryParseArgb(argb, out OfficeColor color)) {
                throw new ArgumentException("Color must be a six-digit RGB or eight-digit ARGB value.", nameof(argb));
            }

            return OfficeColorTransforms.SpreadsheetTint(color, NormalizeSpreadsheetTint(tint)).ToArgbHex();
        }

        private static double NormalizeSpreadsheetTint(double tint) {
            if (double.IsNaN(tint) || double.IsInfinity(tint)) return 0D;
            return Math.Max(-1D, Math.Min(1D, tint));
        }

        internal static string? NormalizeArgb(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            string hex = value!.Trim().TrimStart('#');
            if (hex.Length == 6) {
                hex = "FF" + hex;
            } else if (hex.Length != 8) {
                return null;
            }

            return hex.All(Uri.IsHexDigit) ? hex.ToUpperInvariant() : null;
        }

        private static A.ColorScheme? GetColorScheme(WorkbookPart? workbookPart) =>
            workbookPart?
                .GetPartsOfType<ThemePart>()
                .FirstOrDefault()?
                .Theme?
                .ThemeElements?
                .ColorScheme;

        private static bool TryParseArgb(string value, out OfficeColor color) {
            color = default;
            string? argb = NormalizeArgb(value);
            if (argb == null
                || !byte.TryParse(argb.Substring(0, 2), System.Globalization.NumberStyles.HexNumber, null, out byte alpha)
                || !byte.TryParse(argb.Substring(2, 2), System.Globalization.NumberStyles.HexNumber, null, out byte red)
                || !byte.TryParse(argb.Substring(4, 2), System.Globalization.NumberStyles.HexNumber, null, out byte green)
                || !byte.TryParse(argb.Substring(6, 2), System.Globalization.NumberStyles.HexNumber, null, out byte blue)) {
                return false;
            }

            color = OfficeColor.FromRgba(red, green, blue, alpha);
            return true;
        }
    }
}
