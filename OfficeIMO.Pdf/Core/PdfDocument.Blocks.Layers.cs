namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    private int nextGeneratedLayerId = 1;

    /// <summary>Adds flow content controlled by a generated PDF optional-content layer.</summary>
    internal PdfDocument Layer(string name, Action<PdfContentBuilder> compose, PdfLayerOptions? options = null) {
        Guard.NotNullOrWhiteSpace(name, nameof(name));
        Guard.NotNull(compose, nameof(compose));
        var definition = new PdfLayerDefinition(nextGeneratedLayerId++, name.Trim(), options ?? new PdfLayerOptions());
        var blocks = new List<IPdfBlock>();
        using (PushBlockScope(blocks.Add)) {
            compose(new PdfContentBuilder(this));
        }

        AddBlock(new LayerBlock(definition, blocks));
        return this;
    }
}
