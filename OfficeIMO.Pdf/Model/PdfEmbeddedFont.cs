namespace OfficeIMO.Pdf;

/// <summary>
/// Represents a TrueType font file that should be embedded for a generated standard-font slot.
/// </summary>
public sealed class PdfEmbeddedFont {
    private readonly byte[] _data;

    /// <summary>Creates an embedded-font mapping for a standard PDF font slot.</summary>
    public PdfEmbeddedFont(PdfStandardFont font, byte[] data, string? fontName = null) {
        Guard.StandardFont(font, nameof(font), "PDF embedded font mapping must target one of the supported standard PDF font slots.");
        Guard.NotNull(data, nameof(data));
        if (data.Length == 0) {
            throw new ArgumentException("PDF embedded font data cannot be empty.", nameof(data));
        }

        Font = font;
        _data = (byte[])data.Clone();
        FontName = string.IsNullOrWhiteSpace(fontName) ? null : fontName;
    }

    /// <summary>Standard PDF font slot that this font file replaces in generated output.</summary>
    public PdfStandardFont Font { get; }

    /// <summary>Optional PDF font name override. When null, the TrueType name table is used.</summary>
    public string? FontName { get; }

    /// <summary>TrueType font file bytes.</summary>
    public byte[] Data => (byte[])_data.Clone();

    internal byte[] DataSnapshot => _data;

    internal PdfEmbeddedFont Clone() => this;
}
