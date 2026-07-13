namespace OfficeIMO.Pdf;

/// <summary>
/// Provides externally shaped glyph runs for generated PDF text written with embedded fonts.
/// </summary>
/// <remarks>
/// OfficeIMO.Pdf remains dependency-free; callers can plug in a host-owned shaping engine such as HarfBuzz through this contract.
/// Returning <c>null</c> lets the built-in dependency-free shaping path handle the text.
/// </remarks>
public interface IPdfTextShapingProvider {
    /// <summary>
    /// Shapes text into font glyph identifiers for PDF output.
    /// </summary>
    /// <param name="request">Text, font, and mode information for the run being written.</param>
    /// <returns>A shaped glyph run, or <c>null</c> to fall back to OfficeIMO.Pdf's built-in shaping.</returns>
    PdfTextShapingResult? ShapeText(PdfTextShapingRequest request);
}

/// <summary>
/// Describes a text run and embedded font passed to a PDF text shaping provider.
/// </summary>
public sealed class PdfTextShapingRequest {
    private readonly byte[] _fontData;

    /// <summary>
    /// Creates a shaping request.
    /// </summary>
    /// <param name="text">Original UTF-16 text to shape.</param>
    /// <param name="fontName">PDF font name selected for the text run.</param>
    /// <param name="fontData">Snapshot of the embedded font bytes.</param>
    /// <param name="isOpenTypeCff">True when the font uses OpenType/CFF outlines; false for TrueType outlines.</param>
    /// <param name="fallbackMode">OfficeIMO.Pdf built-in shaping mode configured for fallback handling.</param>
    public PdfTextShapingRequest(string text, string fontName, byte[] fontData, bool isOpenTypeCff, PdfTextShapingMode fallbackMode) {
        Guard.NotNull(text, nameof(text));
        Guard.NotNull(fontData, nameof(fontData));
        Text = text;
        FontName = fontName ?? string.Empty;
        _fontData = fontData.ToArray();
        IsOpenTypeCff = isOpenTypeCff;
        FallbackMode = fallbackMode;
    }

    /// <summary>Original UTF-16 text to shape.</summary>
    public string Text { get; }

    /// <summary>PDF font name selected for the text run.</summary>
    public string FontName { get; }

    /// <summary>Snapshot of the embedded font bytes.</summary>
    public byte[] FontData => (byte[])_fontData.Clone();

    /// <summary>True when the font uses OpenType/CFF outlines; false for TrueType outlines.</summary>
    public bool IsOpenTypeCff { get; }

    /// <summary>OfficeIMO.Pdf built-in shaping mode configured for fallback handling.</summary>
    public PdfTextShapingMode FallbackMode { get; }
}

/// <summary>
/// A shaped glyph run returned by an external PDF text shaping provider.
/// </summary>
public sealed class PdfTextShapingResult {
    /// <summary>
    /// Creates a shaping result from glyph mappings.
    /// </summary>
    /// <param name="glyphs">Glyph identifiers and source-text mappings in visual write order.</param>
    public PdfTextShapingResult(IEnumerable<PdfShapedGlyph> glyphs) {
        Guard.NotNull(glyphs, nameof(glyphs));
        Glyphs = glyphs.ToArray();
    }

    /// <summary>Glyph identifiers and source-text mappings in visual write order.</summary>
    public IReadOnlyList<PdfShapedGlyph> Glyphs { get; }
}

/// <summary>
/// Maps one shaped font glyph back to the source text it represents.
/// </summary>
public readonly struct PdfShapedGlyph {
    /// <summary>
    /// Creates a shaped glyph mapping.
    /// </summary>
    /// <param name="glyphId">Font glyph identifier to write into the PDF content stream.</param>
    /// <param name="unicodeText">Original Unicode text represented by the shaped glyph for ToUnicode extraction.</param>
    /// <param name="textIndex">UTF-16 index in the original text where this glyph's source text begins.</param>
    public PdfShapedGlyph(int glyphId, string unicodeText, int textIndex) {
        if (glyphId <= 0) {
            throw new ArgumentOutOfRangeException(nameof(glyphId), "Shaped PDF glyph identifiers must be positive.");
        }

        GlyphId = glyphId;
        Guard.NotNull(unicodeText, nameof(unicodeText));
        UnicodeText = unicodeText;
        TextIndex = textIndex;
    }

    /// <summary>Font glyph identifier to write into the PDF content stream.</summary>
    public int GlyphId { get; }

    /// <summary>Original Unicode text represented by the shaped glyph for ToUnicode extraction.</summary>
    public string UnicodeText { get; }

    /// <summary>UTF-16 index in the original text where this glyph's source text begins.</summary>
    public int TextIndex { get; }
}
