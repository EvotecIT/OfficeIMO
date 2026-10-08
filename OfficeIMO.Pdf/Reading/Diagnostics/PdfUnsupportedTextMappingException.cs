namespace OfficeIMO.Pdf;

/// <summary>Identifies composite text that cannot be decoded through a usable Unicode mapping.</summary>
internal sealed class PdfUnsupportedTextMappingException : NotSupportedException {
    internal PdfUnsupportedTextMappingException(PdfFontResource font)
        : base("Composite font /" + font.ResourceName + " uses " + font.Encoding +
            " without a usable ToUnicode mapping. Its text cannot be extracted or searched reliably.") {
    }
}
