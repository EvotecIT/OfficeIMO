using OfficeIMO.Excel.Utilities;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>Resolves native theme slots against this workbook and status roles against its creation palette.</summary>
        internal OfficeColor? ResolveColor(OfficeColor color) {
            uint? index = OfficeOpenXmlThemeColorResolver.GetSpreadsheetThemeColorIndex(color.ThemeSlot);
            return index.HasValue
                ? ExcelColor.FromSpreadsheetColor(null, ResolveThemeColorArgb(index.Value))
                : color.Resolve(EffectiveCreationTheme);
        }

        /// <summary>
        /// Resolves a SpreadsheetML theme color to an ARGB hexadecimal value.
        /// </summary>
        /// <param name="themeIndex">
        /// Zero-based SpreadsheetML theme index: light 1, dark 1, light 2, dark 2,
        /// followed by accent colors 1 through 6.
        /// </param>
        /// <param name="tint">
        /// Optional SpreadsheetML tint between -1 and 1.
        /// </param>
        /// <returns>An eight-character ARGB value, or <see langword="null"/> when the theme color cannot be resolved.</returns>
        public string? ResolveThemeColorArgb(uint themeIndex, double? tint = null) =>
            ExcelThemeColorResolver.ResolveTheme(themeIndex, tint, WorkbookPartRoot);
    }
}
