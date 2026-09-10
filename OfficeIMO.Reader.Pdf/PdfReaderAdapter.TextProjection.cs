using OfficeIMO.Pdf;

namespace OfficeIMO.Reader.Pdf;

internal static partial class PdfReaderAdapter {
    private static string GetChunkSourceKind(IReadOnlyList<PdfLogicalPage> pages, string containerKind) {
        bool hasTables = pages.Any(page => page.Tables.Count > 0);
        bool hasText = pages.Any(page => page.TextBlocks.Any(block => !string.IsNullOrWhiteSpace(block.Text)));
        if (!hasTables && !hasText) return "visual";
        // Classify table-only text at its source. A downstream partial read result
        // cannot infer this safely merely from the presence of a table elsewhere.
        bool hasIndependentText = pages.Any(page =>
            page.TextBlocks.Any(block => !block.IsTableContent && !string.IsNullOrWhiteSpace(block.Text))
            || page.Elements.OfType<PdfLogicalLeaderRow>().Any()
            || page.Links.Count > 0 || page.FormWidgets.Count > 0);
        return hasTables && !hasIndependentText ? "table" : containerKind;
    }
}
