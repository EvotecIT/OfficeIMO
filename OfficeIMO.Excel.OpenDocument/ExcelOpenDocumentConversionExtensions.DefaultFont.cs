using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument;

namespace OfficeIMO.Excel.OpenDocument;

public static partial class ExcelOpenDocumentConversionExtensions {
    // Workbook font zero defines column units independently of individually styled cells.
    private static void ExportWorkbookDefaultFont(ExcelDocument source, OdsDocument target) {
        if (source.GetWorkbookDefaultFontInfo() is not OfficeFontInfo font) return;
        OdfStyle style = target.Styles.GetOrCreateDefault(OdfStyleFamily.TableCell);
        style.FontFamily = font.FamilyName;
        style.FontSize = OdfLength.Points(font.Size);
        style.Bold = font.IsBold;
        style.Italic = font.IsItalic;
    }

    private static void ImportWorkbookDefaultFont(OdsDocument source, ExcelDocument target) {
        OdfStyle? style = source.Styles.FindDefault(OdfStyleFamily.TableCell);
        if (style == null) return;
        target.EnsureWorkbookThemeAndStyles();
        Font? font = target.OpenXmlDocument.WorkbookPart?.WorkbookStylesPart?.Stylesheet?.Fonts?.Elements<Font>().FirstOrDefault();
        if (font == null) return;
        if (OdfFontFamilySyntax.TryParse(style.FontFamily, out OdfFontFamilySyntax? family))
            font.FontName = new FontName { Val = family!.PrimaryFamily };
        if (style.FontSize.HasValue) font.FontSize = new FontSize { Val = OdfTextStyleResolver.ResolveFontSize(source.Styles.Resolve(style)) };
        if (style.Bold.HasValue) font.Bold = new Bold { Val = style.Bold.Value };
        if (style.Italic.HasValue) font.Italic = new Italic { Val = style.Italic.Value };
    }
}
