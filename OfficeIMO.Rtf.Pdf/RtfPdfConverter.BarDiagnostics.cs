using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

internal static partial class RtfPdfConverter {
    /// <summary>Reports the selected projection's bar fidelity once per source paragraph.</summary>
    private static void RecordParagraphBarTabDiagnostic(RtfDocument document, RtfParagraph paragraph,
        RtfToPdfOptions options, string? omittedSource = null) {
        if (!document.GetParagraphFormatting(paragraph).TabStops.Any(stop => stop.Alignment == RtfTabAlignment.Bar) ||
            !options.DiagnosedBarTabParagraphs.Add(paragraph)) return;
        if (omittedSource != null) {
            options.Report.Add(new PdfCore.PdfConversionWarning(
                "OfficeIMO.Rtf.Pdf", "BarTabOmitted", omittedSource,
                "RTF bar tabs are omitted when paragraph content is projected as plain text.",
                PdfCore.PdfConversionWarningSeverity.Warning, OfficeConversionLossKind.Omission,
                details: new Dictionary<string, string> { ["RtfAction"] = RtfConversionAction.Omitted.ToString() }));
            return;
        }
        AddConversionWarning(options, "BarTabAppearanceApproximated", "Paragraph/Tabs",
            "RTF bar tabs use the shared PDF engine's default stroke appearance; native RTF pen width and line extents are not independently qualified.",
            RtfConversionAction.Flattened);
    }
}
