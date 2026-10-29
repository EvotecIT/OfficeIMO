using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    private static PdfCore.PdfTextRun? CreateNativeFootnoteSeparatorRule(OpenXmlElement mark,
        WordParagraph run, NativeDocumentDefaults defaults, NativeFontMap? fonts,
        NativeTableRunStyleDefaults table, NativeNoteNumbering? notes) {
        // A separator keeps its source run's font metrics and strike band,
        // including the selected face and baseline shift. The Word paragraph
        // style also treats this special character as metric-bearing content.
        PdfCore.PdfTextRun metricStyle = CreateNativeCellTextRun(" ", run,
            NativeTableStyleDefaults.Empty with { RunStyle = table }, defaults, fonts);
        PdfCore.PdfWriter.TextDecorationMeasurement? band = fonts?.MeasureStrikethrough(metricStyle,
            NativeLayoutDevicePointsPerPixel);
        if (band == null) return null;
        // Native's short special character has a two-inch advance (1200 pixels
        // at its 600-dpi layout device). Both marks clamp at the real remaining
        // line frame; the continuation mark has no short-rule maximum.
        double? maximum = mark is W.SeparatorMark ? 2D * 72D : null;
        var rule = new PdfCore.PdfInlineRule(band.Thickness, metricStyle.Color ?? global::OfficeIMO.OfficeColor.Black,
            maximumWidth: maximum, baselineOffset: band.Position - band.Thickness);
        if (metricStyle.Underline) {
            PdfCore.PdfWriter.TextDecorationMeasurement? underline = fonts?.MeasureUnderline(metricStyle,
                NativeLayoutDevicePointsPerPixel);
            if (metricStyle.UnderlineStyle == OfficeTextDecorationStyle.Single && underline != null) {
                rule = rule.WithAdditionalBand(underline.Thickness,
                    metricStyle.DecorationColor ?? metricStyle.Color ?? global::OfficeIMO.OfficeColor.Black,
                    underline.Position - underline.Thickness);
            } else {
                notes?.ReportSeparatorDecorationOmission("The separator's underline pattern or font underline metrics are outside the qualified rule profile.");
            }
        }
        if (metricStyle.StrikeStyle is not (OfficeTextDecorationStyle.None or OfficeTextDecorationStyle.Single))
            notes?.ReportSeparatorDecorationOmission("The separator's additional strikethrough pattern is outside the qualified rule profile.");
        return PdfCore.PdfTextRun.Inline(rule, metricStyle);
    }
}
