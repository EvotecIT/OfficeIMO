using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Resolves the built-in Normal style's font through its style-format and
        /// font indexes. Minimal workbooks without that style use font zero.
        /// Column character units use this font independently of cell formatting.
        /// </summary>
        internal OfficeFontInfo? GetNormalStyleFontInfo() {
            Stylesheet? styles = WorkbookPartRoot.WorkbookStylesPart?.Stylesheet;
            CellStyle? normal = styles?.CellStyles?.Elements<CellStyle>()
                .FirstOrDefault(style => style.BuiltinId?.Value == 0U);
            uint formatIndex = normal?.FormatId?.Value ?? 0U;
            CellFormat? format = formatIndex <= int.MaxValue
                ? styles?.CellStyleFormats?.Elements<CellFormat>().ElementAtOrDefault((int)formatIndex)
                : null;
            uint fontIndex = normal == null ? 0U : format?.FontId?.Value ?? 0U;
            Font? font = fontIndex <= int.MaxValue
                ? styles?.Fonts?.Elements<Font>().ElementAtOrDefault((int)fontIndex)
                : null;
            string? name = font?.GetFirstChild<FontName>()?.Val?.Value;
            if (string.IsNullOrWhiteSpace(name) || font == null) return null;

            OfficeFontStyle face = OfficeFontStyle.Regular;
            if (ExcelOpenXmlFontProperty.IsEnabled(font.GetFirstChild<Bold>())) face |= OfficeFontStyle.Bold;
            if (ExcelOpenXmlFontProperty.IsEnabled(font.GetFirstChild<Italic>())) face |= OfficeFontStyle.Italic;
            if (ExcelOpenXmlFontProperty.IsUnderlineEnabled(font.GetFirstChild<Underline>())) face |= OfficeFontStyle.Underline;
            return new OfficeFontInfo(name, font.GetFirstChild<FontSize>()?.Val?.Value ?? 11D, face);
        }
    }
}
