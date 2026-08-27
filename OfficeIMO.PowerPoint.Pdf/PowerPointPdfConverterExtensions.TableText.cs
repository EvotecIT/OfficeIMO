using DocumentFormat.OpenXml.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    private static PdfCore.PdfTextRun CreatePdfTableCellTextRun(PptCore.PowerPointTableCell cell, A.Run run, string text, string? fallbackFontFamily) {
        A.RunProperties? properties = run.RunProperties;
        string? fontFamily = ReadRunFontName(properties) ?? cell.FontName ?? fallbackFontFamily;
        A.TextUnderlineValues? underline = properties?.Underline?.Value;
        A.TextStrikeValues? strike = properties?.Strike?.Value;
        int? baseline = properties?.Baseline?.Value;
        A.TextCapsValues? capitalization = properties?.Capital?.Value;
        if (capitalization == A.TextCapsValues.All || capitalization == A.TextCapsValues.Small) {
            text = OfficeIMO.Drawing.OfficeTextCaseTransformer.Apply(text, OfficeIMO.Drawing.OfficeTextCase.Uppercase, System.Globalization.CultureInfo.InvariantCulture);
        }
        return new PdfCore.PdfTextRun(
            text,
            bold: properties?.Bold?.Value ?? cell.Bold,
            underline: underline.HasValue && underline.Value != TextUnderlineValues.None,
            color: ParsePdfColor(ReadRunColor(properties) ?? cell.Color),
            italic: properties?.Italic?.Value ?? cell.Italic,
            strike: strike.HasValue && strike.Value != A.TextStrikeValues.NoStrike,
            fontSize: ReadRunFontSize(properties) ?? cell.FontSize,
            font: MapFont(fontFamily),
            baseline: baseline > 0 ? PdfCore.PdfTextBaseline.Superscript : baseline < 0 ? PdfCore.PdfTextBaseline.Subscript : PdfCore.PdfTextBaseline.Normal,
            fontFamily: fontFamily,
            underlineStyle: MapPowerPointUnderline(underline),
            strikeStyle: strike == A.TextStrikeValues.DoubleStrike
                ? OfficeIMO.Drawing.OfficeTextDecorationStyle.Double
                : strike == A.TextStrikeValues.SingleStrike
                    ? OfficeIMO.Drawing.OfficeTextDecorationStyle.Single
                    : OfficeIMO.Drawing.OfficeTextDecorationStyle.None);
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
