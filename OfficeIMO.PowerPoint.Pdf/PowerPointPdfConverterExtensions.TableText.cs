using DocumentFormat.OpenXml.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    private static PdfCore.TextRun CreatePdfTableCellTextRun(PptCore.PowerPointTableCell cell, A.Run run, string text, string? fallbackFontFamily) {
        A.RunProperties? properties = run.RunProperties;
        string? fontFamily = ReadRunFontName(properties) ?? cell.FontName ?? fallbackFontFamily;
        return new PdfCore.TextRun(
            text,
            bold: properties?.Bold?.Value ?? cell.Bold,
            underline: properties?.Underline?.Value == TextUnderlineValues.Single,
            color: ParsePdfColor(ReadRunColor(properties) ?? cell.Color),
            italic: properties?.Italic?.Value ?? cell.Italic,
            fontSize: ReadRunFontSize(properties) ?? cell.FontSize,
            font: MapFont(fontFamily),
            fontFamily: fontFamily);
    }

    private static string? ReadRunColor(A.RunProperties? properties) =>
        properties?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val;

    private static double? ReadRunFontSize(A.RunProperties? properties) {
        int? size = properties?.FontSize?.Value;
        return size.HasValue ? size.Value / 100D : null;
    }

    private static string? ReadRunFontName(A.RunProperties? properties) =>
        properties?.GetFirstChild<A.LatinFont>()?.Typeface;
}
