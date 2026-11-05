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

    private static string BuildPlainText(IReadOnlyList<PdfLogicalPage> pages) =>
        string.Join(Environment.NewLine + Environment.NewLine, pages.Select(PdfDocumentReadResult.GetCanonicalPageText));

    private static IEnumerable<ReaderChunk> BuildChunksFromText(string markdown, string text, SourceMetadata source,
        ReaderOptions readerOptions, int? page, int sourceBlockIndex, string blockKind, string blockAnchor,
        IReadOnlyList<ReaderTable>? tables, IReadOnlyList<ReaderVisual>? visuals, IReadOnlyList<ReaderFormField>? formFields,
        IReadOnlyList<ReaderActionSummary>? actions, ReaderChunkDiagnostics diagnostics, string idPrefix, int maxChars,
        CancellationToken cancellationToken) {
        var markdownParts = SplitText(markdown, maxChars);
        IReadOnlyList<string> textParts = DocumentReaderEngine.SplitAdapterProjection(text, maxChars);
        if (markdownParts.Count == 0 && textParts.Count == 0) {
            blockKind = "warning";
            string warning = page.HasValue
                ? "PDF page " + page.Value.ToString(CultureInfo.InvariantCulture) + " produced no readable text."
                : "PDF content produced no readable text.";
            markdownParts = new[] { new TextPart(warning, new[] { warning }) };
            textParts = new[] { warning };
        }

        int count = Math.Max(markdownParts.Count, textParts.Count);
        for (int i = 0; i < count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            TextPart? markdownPart = i < markdownParts.Count ? markdownParts[i] : null;
            yield return EnrichChunk(new ReaderChunk {
                Id = count == 1 ? idPrefix : idPrefix + "-part-" + i.ToString("D4", CultureInfo.InvariantCulture),
                Kind = ReaderInputKind.Pdf,
                Location = new ReaderLocation {
                    Path = source.Path,
                    SourceBlockIndex = sourceBlockIndex,
                    SourceBlockKind = blockKind,
                    BlockAnchor = count == 1 ? blockAnchor : blockAnchor + "-part-" + i.ToString("D4", CultureInfo.InvariantCulture),
                    Page = page
                },
                Text = i < textParts.Count ? textParts[i] : string.Empty,
                Markdown = markdownPart?.Text,
                ContinuesPreviousChunk = i > 0,
                Tables = i == 0 ? tables : null,
                Visuals = i == 0 ? visuals : null,
                FormFields = i == 0 ? formFields : null,
                Actions = i == 0 ? actions : null,
                Diagnostics = diagnostics,
                Warnings = markdownPart?.Warnings
            }, source, readerOptions.ComputeHashes);
        }
    }
}
