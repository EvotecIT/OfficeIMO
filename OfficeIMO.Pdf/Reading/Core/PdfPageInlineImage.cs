namespace OfficeIMO.Pdf;

internal sealed class PdfPageInlineImage {
    public PdfPageInlineImage(string resourceName, PdfStream stream) {
        ResourceName = resourceName;
        Stream = stream;
        DirectStreamIdentity = PdfDirectStreamIdentity.Compute(stream);
    }

    public string ResourceName { get; }

    public PdfStream Stream { get; }

    public int DirectStreamIdentity { get; }
}
