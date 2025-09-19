namespace OfficeIMO.Pdf;

internal sealed class PdfFontResource {
    public string ResourceName { get; }
    public string BaseFont { get; }
    public string Encoding { get; }
    public bool HasToUnicode { get; }
    public PdfFontResource(string resourceName, string baseFont, string encoding, bool hasToUnicode) {
        ResourceName = resourceName; BaseFont = baseFont; Encoding = encoding; HasToUnicode = hasToUnicode;
    }
}

