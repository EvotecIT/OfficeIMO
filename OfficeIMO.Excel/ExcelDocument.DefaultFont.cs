using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Returns workbook font zero, which native worksheet column units use
        /// independently of cell formatting and named style-format references.
        /// </summary>
        internal OfficeFontInfo? GetWorkbookDefaultFontInfo() {
            Stylesheet? styles = WorkbookPartRoot.WorkbookStylesPart?.Stylesheet;
            Font? font = styles?.Fonts?.Elements<Font>().FirstOrDefault();
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
