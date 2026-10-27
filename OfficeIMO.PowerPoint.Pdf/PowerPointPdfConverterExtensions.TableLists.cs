using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    private static PdfCore.PdfTextRun CreatePdfTableCellListMarker(
        PptCore.PowerPointTableCell cell, PptCore.PowerPointParagraph paragraph, string marker,
        string? fallbackFontFamily, double fallbackFontSize, int slideNumber, PowerPointPdfExportOptions options,
        System.Collections.Generic.IReadOnlyList<A.TextCharacterPropertiesType> tableDefaults) {
        A.RunProperties? properties = paragraph.Paragraph.Elements<A.Run>().FirstOrDefault()?.RunProperties;
        PdfCore.PdfTextRun inherited = CreatePdfTableCellTextRun(cell, paragraph.Paragraph,
            properties, marker, fallbackFontFamily, slideNumber, options, tableDefaults, fallbackFontSize);
        string? family = paragraph.BulletFontName ?? inherited.FontFamily;
        double? size = paragraph.BulletSizePoints;
        if (!size.HasValue && paragraph.BulletSizePercent is int percent)
            size = (inherited.FontSize ?? fallbackFontSize) * percent / 100d;
        return new PdfCore.PdfTextRun(marker, bold: inherited.Bold, italic: inherited.Italic,
            color: inherited.Color, fontSize: size ?? inherited.FontSize,
            font: MapFont(family), fontFamily: family);
    }
}
