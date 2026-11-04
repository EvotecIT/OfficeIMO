using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel.Utilities {
    /// <summary>Resolves worksheet column dimensions in the 96-DPI drawing coordinate system.</summary>
    internal static class ExcelColumnGeometry {
        /// <summary>
        /// Uses the workbook's font-zero digit advances from supplied export faces, or an installed face when no collection is supplied.
        /// Grid dimensions use whole-pixel digit widths; missing faces retain the seven-pixel fallback.
        /// </summary>
        internal static double GetMaximumDigitWidth(WorksheetPart? worksheetPart, OfficeFontFaceCollection? fonts = null) => GetMaximumDigitWidth(worksheetPart, out _, fonts);

        internal static double GetMaximumDigitWidth(WorksheetPart? worksheetPart, out bool resolved, OfficeFontFaceCollection? fonts = null) {
            resolved = false;
            if (fonts?.Faces.Count == 0) fonts = null;
            WorkbookPart? workbook = worksheetPart?.GetParentParts().OfType<WorkbookPart>().FirstOrDefault();
            Font? font = workbook?.WorkbookStylesPart?.Stylesheet?.Fonts?.Elements<Font>().FirstOrDefault();
            string? family = font?.GetFirstChild<FontName>()?.Val?.Value;
            if (font == null || string.IsNullOrWhiteSpace(family)) return 7D;

            OfficeFontStyle style = OfficeFontStyle.Regular;
            if (ExcelOpenXmlFontProperty.IsEnabled(font.GetFirstChild<Bold>())) style |= OfficeFontStyle.Bold;
            if (ExcelOpenXmlFontProperty.IsEnabled(font.GetFirstChild<Italic>())) style |= OfficeFontStyle.Italic;
            double size = font.GetFirstChild<FontSize>()?.Val?.Value ?? 11D;
            if (size <= 0D || double.IsNaN(size) || double.IsInfinity(size)) return 7D;
            OfficeTrueTypeFont? program = null;
            if (fonts == null) {
                program = OfficeTrueTypeFont.TryLoadFontFamilyForText(family, style, "0123456789", out _);
                if (program == null || !program.MatchesFamilyName(family!)) return 7D;
            }
            double maximum = 0D;
            for (char digit = '0'; digit <= '9'; digit++) {
                double width;
                if (fonts != null) {
                    if (!fonts.TryMeasureText(digit.ToString(), size * 96D / 72D, family, style, out width)) return 7D;
                } else width = program!.Measure(digit.ToString(), size * 96D / 72D);
                maximum = Math.Max(maximum, width);
            }
            resolved = maximum > 0D && !double.IsNaN(maximum) && !double.IsInfinity(maximum);
            return maximum > 0D && !double.IsNaN(maximum) && !double.IsInfinity(maximum)
                ? Math.Max(1D, Math.Round(maximum, MidpointRounding.AwayFromZero)) : 7D;
        }

        /// <summary>
        /// Explicit stored widths already include cell padding. Implicit base widths exclude
        /// five pixels of padding and round upward to the next eight-pixel default-column boundary.
        /// </summary>
        internal static int GetWidthPixels(Column? column, SheetFormatProperties? format, double maximumDigitWidth, bool includeHidden = false) {
            if (!includeHidden && column?.Hidden?.Value == true) return 0;
            double? stored = column?.Width?.Value > 0D
                ? column.Width.Value : format?.DefaultColumnWidth?.Value;
            double pixels;
            if (stored is double width && width > 0D) {
                pixels = Math.Truncate((256D * width + Math.Truncate(128D / maximumDigitWidth)) / 256D * maximumDigitWidth);
            } else {
                double baseWidth = format?.BaseColumnWidth?.Value > 0U ? format.BaseColumnWidth.Value : 8D;
                pixels = Math.Ceiling((baseWidth * maximumDigitWidth + 5D) / 8D) * 8D;
            }
            return Math.Max(1, (int)Math.Min(int.MaxValue, pixels));
        }
    }
}
