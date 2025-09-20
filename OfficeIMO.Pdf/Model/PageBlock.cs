namespace OfficeIMO.Pdf;

internal sealed class PageBlock : IPdfBlock {
    public PdfOptions Options { get; }
    public System.Collections.Generic.List<IPdfBlock> Blocks { get; } = new();
    public PageBlock(PdfOptions options) {
        Guard.NotNull(options, nameof(options));
        Options = options;
    }
}
