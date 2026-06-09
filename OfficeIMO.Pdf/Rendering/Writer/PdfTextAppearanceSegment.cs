namespace OfficeIMO.Pdf;

internal sealed class PdfTextAppearanceSegment {
    public PdfTextAppearanceSegment(string fontResourceName, string encodedHex) {
        Guard.NotNullOrWhiteSpace(fontResourceName, nameof(fontResourceName));
        Guard.NotNull(encodedHex, nameof(encodedHex));

        FontResourceName = fontResourceName;
        EncodedHex = encodedHex;
    }

    public string FontResourceName { get; }

    public string EncodedHex { get; }
}
