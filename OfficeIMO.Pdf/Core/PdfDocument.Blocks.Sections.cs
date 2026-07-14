namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    private int nextGeneratedSectionId = 1;

    /// <summary>Adds a semantic section backed by an outline entry and named destination.</summary>
    public PdfDocument Section(
        string title,
        Action<PdfItemCompose> compose,
        PdfSectionOptions? options = null) {
        Guard.NotNullOrWhiteSpace(title, nameof(title));
        Guard.NotNull(compose, nameof(compose));
        PdfSectionOptions requested = options ?? new PdfSectionOptions();
        string destinationName = string.IsNullOrWhiteSpace(requested.DestinationName)
            ? "section-" + nextGeneratedSectionId++.ToString(System.Globalization.CultureInfo.InvariantCulture)
            : requested.DestinationName!.Trim();
        var blocks = new List<IPdfBlock>();
        using (PushBlockScope(blocks.Add)) {
            compose(new PdfItemCompose(this));
        }

        AddBlock(new SectionBlock(title, blocks, requested.Clone(destinationName)));
        return this;
    }

    /// <summary>Adds an automatically generated, internally linked table of contents.</summary>
    public PdfDocument TableOfContents(PdfTableOfContentsOptions? options = null) {
        AddBlock(new TableOfContentsBlock(options));
        return this;
    }
}
