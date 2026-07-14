namespace OfficeIMO.Pdf;

internal sealed class TableOfContentsBlock : IPdfBlock {
    public TableOfContentsBlock(PdfTableOfContentsOptions? options) {
        Options = options?.Clone() ?? new PdfTableOfContentsOptions();
    }

    public PdfTableOfContentsOptions Options { get; }
}
