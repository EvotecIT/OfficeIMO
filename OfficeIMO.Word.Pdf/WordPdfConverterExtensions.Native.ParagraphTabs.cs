using System.Collections.Generic;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>
    /// Projects effective Word tabs into the resolved paragraph text frame for
    /// body, note, running-story and table-cell paragraphs without changing source XML.
    /// </summary>
    private static IReadOnlyList<PdfCore.PdfTabStop> ResolveNativeParagraphTabStops(
        WordParagraph paragraph, double leftIndent, double firstLineIndent) {
        var result = new List<PdfCore.PdfTabStop>();
        foreach (WordTabStop tabStop in GetNativeParagraphEffectiveTabStops(paragraph)
            .Where(tabStop => tabStop.Alignment.ToOpenXml() != DocumentFormat.OpenXml.Wordprocessing.TabStopValues.Clear)
            .OrderBy(tabStop => tabStop.Position)) {
            // Source stops are relative to the containing text margin, while
            // PDF stops use the continuation-line text frame. A hanging first
            // line can reach stops at or before that frame's origin.
            double framePosition = tabStop.Position.ToPoints() - leftIndent;
            result.Add(new PdfCore.PdfTabStop(
                framePosition,
                MapNativeTabAlignment(tabStop.Alignment),
                MapNativeTabLeader(tabStop.Leader)) {
                    DecimalSeparator = GetNativeDecimalTabSeparator(paragraph),
                    // Native Word emits this explicit pen width independently of
                    // the paragraph font. The shared engine owns line geometry.
                    BarWidth = 0.14D
                });
        }

        // The existing compatibility policy also governs Word's implicit
        // hanging-indent stop outside numbered paragraphs.
        if (firstLineIndent < 0D && !IgnoresNativeNumberingIndentStop(paragraph) &&
            !result.Any(tabStop => tabStop.Alignment != PdfCore.PdfTabAlignment.Bar && tabStop.Position == 0D)) {
            result.Add(new PdfCore.PdfTabStop(0D));
        }

        return result;
    }

    private static void RefreshNativeParagraphTabStops(WordParagraph paragraph, PdfCore.PdfParagraphStyle style) {
        style.TabStops.Clear();
        foreach (PdfCore.PdfTabStop stop in ResolveNativeParagraphTabStops(paragraph, style.LeftIndent, style.FirstLineIndent))
            style.TabStops.Add(stop);
    }

    private static bool HasNativeParagraphBarTabs(WordParagraph paragraph) =>
        GetNativeParagraphEffectiveTabStops(paragraph).Any(stop =>
            stop.Alignment.ToOpenXml() == DocumentFormat.OpenXml.Wordprocessing.TabStopValues.Bar);
}
