namespace OfficeIMO.Pdf;

internal sealed class SectionBlock : IPdfBlock {
    public SectionBlock(string title, IEnumerable<IPdfBlock> blocks, PdfSectionOptions options) {
        Guard.NotNullOrWhiteSpace(title, nameof(title));
        Guard.NotNull(blocks, nameof(blocks));
        Guard.NotNull(options, nameof(options));
        Title = title;
        Blocks = Array.AsReadOnly(blocks.ToArray());
        Options = options;
    }

    public string Title { get; }
    public IReadOnlyList<IPdfBlock> Blocks { get; }
    public PdfSectionOptions Options { get; }
    public string DestinationName => Options.DestinationName!;
}
