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
            .Where(tabStop => IsNativeRenderableTextTabStop(tabStop.Alignment))
            .OrderBy(tabStop => tabStop.Position)) {
            // Source stops are relative to the containing text margin, while
            // PDF stops use the continuation-line text frame. A hanging first
            // line can reach stops at or before that frame's origin.
            double framePosition = tabStop.Position / 20D - leftIndent;
            result.Add(new PdfCore.PdfTabStop(
                framePosition,
                MapNativeTabAlignment(tabStop.Alignment),
                MapNativeTabLeader(tabStop.Leader)) { DecimalSeparator = GetNativeDecimalTabSeparator(paragraph) });
        }

        // The existing compatibility policy also governs Word's implicit
        // hanging-indent stop outside numbered paragraphs.
        if (firstLineIndent < 0D && !IgnoresNativeNumberingIndentStop(paragraph) &&
            !result.Any(tabStop => tabStop.Position == 0D)) {
            result.Add(new PdfCore.PdfTabStop(0D));
        }

        return result;
    }
}
