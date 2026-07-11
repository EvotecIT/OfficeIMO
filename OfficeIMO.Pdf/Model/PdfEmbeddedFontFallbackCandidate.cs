namespace OfficeIMO.Pdf;

/// <summary>
/// Describes one embedded font candidate used when planning generated PDF text fallback.
/// </summary>
public sealed class PdfEmbeddedFontFallbackCandidate {
    private readonly byte[] _fontData;

    /// <summary>
    /// Creates a fallback candidate from TrueType or OpenType/CFF font bytes.
    /// </summary>
    /// <param name="fontName">Display name used in fallback segments and diagnostics.</param>
    /// <param name="trueTypeFont">TrueType or OpenType/CFF font bytes to inspect for Unicode glyph coverage.</param>
    public PdfEmbeddedFontFallbackCandidate(string fontName, byte[] trueTypeFont) {
        Guard.NotNullOrWhiteSpace(fontName, nameof(fontName));
        Guard.NotNull(trueTypeFont, nameof(trueTypeFont));
        if (trueTypeFont.Length == 0) {
            throw new ArgumentException("Embedded font fallback data cannot be empty.", nameof(trueTypeFont));
        }

        FontName = fontName;
        _fontData = trueTypeFont.ToArray();
    }

    /// <summary>Display name used in fallback segments and diagnostics.</summary>
    public string FontName { get; }

    internal byte[] DataSnapshot => _fontData;
}
