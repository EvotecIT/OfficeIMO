namespace OfficeIMO.Pdf;

/// <summary>Identifies shown text that cannot be decoded under the configured font-mapping policy.</summary>
public sealed class PdfUnsupportedTextMappingException : NotSupportedException {
    internal PdfUnsupportedTextMappingException(PdfFontResource font)
        : base("Font /" + font.ResourceName + " uses " + font.Encoding +
            " without a usable mapping for the shown character codes under the configured decoding policy. Its text cannot be extracted or searched reliably.") {
        FontResource = font.ResourceName;
        BaseFont = font.BaseFont;
        EncodingName = font.Encoding;
        HasDeclaredToUnicode = font.HasToUnicode;
    }

    /// <summary>Font resource whose shown character codes could not be mapped.</summary>
    public string FontResource { get; }
    /// <summary>Declared PDF base font name.</summary>
    public string BaseFont { get; }
    /// <summary>Declared composite encoding.</summary>
    public string EncodingName { get; }
    /// <summary>Whether a ToUnicode entry was declared, although unusable for the shown codes.</summary>
    public bool HasDeclaredToUnicode { get; }
}
