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
    public PdfTextShapingRequest(string text, string fontName, byte[] fontData, bool isOpenTypeCff, PdfTextShapingMode fallbackMode)
        : this(text, fontName, fontData, isOpenTypeCff, fallbackMode, unitsPerEm: 1000, PdfTextDirection.Auto, language: null) {
    }

    /// <summary>
    /// Creates a shaping request with font metrics and text-context hints.
    /// </summary>
    /// <param name="text">Original UTF-16 text to shape.</param>
    /// <param name="fontName">PDF font name selected for the text run.</param>
    /// <param name="fontData">Snapshot of the embedded font bytes.</param>
    /// <param name="isOpenTypeCff">True when the font uses OpenType/CFF outlines; false for TrueType outlines.</param>
    /// <param name="fallbackMode">OfficeIMO.Pdf built-in shaping mode configured for fallback handling.</param>
    /// <param name="unitsPerEm">Design units per em used by the supplied font.</param>
    /// <param name="direction">Resolved base direction hint for the run.</param>
    /// <param name="language">Optional BCP 47 document language hint.</param>
    public PdfTextShapingRequest(string text, string fontName, byte[] fontData, bool isOpenTypeCff, PdfTextShapingMode fallbackMode, int unitsPerEm, PdfTextDirection direction, string? language) {
        Guard.NotNull(text, nameof(text));
        Guard.NotNull(fontData, nameof(fontData));
        if (unitsPerEm <= 0) {
            throw new ArgumentOutOfRangeException(nameof(unitsPerEm), "Text shaping units per em must be positive.");
        }

        if (direction != PdfTextDirection.Auto && direction != PdfTextDirection.LeftToRight && direction != PdfTextDirection.RightToLeft) {
            throw new ArgumentOutOfRangeException(nameof(direction), "Text shaping direction must be Auto, LeftToRight, or RightToLeft.");
        }

        Text = text;
        FontName = fontName ?? string.Empty;
        _fontData = fontData.ToArray();
        IsOpenTypeCff = isOpenTypeCff;
        FallbackMode = fallbackMode;
        UnitsPerEm = unitsPerEm;
        Direction = direction;
        Language = string.IsNullOrWhiteSpace(language) ? null : language;
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

    /// <summary>Design units per em used by glyph advances and offsets returned for this font.</summary>
    public int UnitsPerEm { get; }

    /// <summary>Base direction inferred from the first strong character in the run.</summary>
    public PdfTextDirection Direction { get; }

    /// <summary>Optional BCP 47 document language hint.</summary>
    public string? Language { get; }
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
    public PdfShapedGlyph(int glyphId, string unicodeText, int textIndex)
        : this(glyphId, unicodeText, textIndex, advanceWidth: null, offsetX: 0, offsetY: 0) {
    }

    /// <summary>
    /// Creates a positioned shaped glyph in the font design units declared by <see cref="PdfTextShapingRequest.UnitsPerEm"/>.
    /// </summary>
    /// <param name="glyphId">Font glyph identifier to write into the PDF content stream.</param>
    /// <param name="unicodeText">Original Unicode text represented by the shaped glyph for ToUnicode extraction.</param>
    /// <param name="textIndex">UTF-16 index in the original text where this glyph's source text begins.</param>
    /// <param name="advanceWidth">Horizontal shaped advance in font design units.</param>
    /// <param name="offsetX">Horizontal glyph placement offset in font design units.</param>
    /// <param name="offsetY">Vertical glyph placement offset in font design units.</param>
    public PdfShapedGlyph(int glyphId, string unicodeText, int textIndex, int advanceWidth, int offsetX = 0, int offsetY = 0)
        : this(glyphId, unicodeText, textIndex, (int?)advanceWidth, offsetX, offsetY) {
    }

    private PdfShapedGlyph(int glyphId, string unicodeText, int textIndex, int? advanceWidth, int offsetX, int offsetY) {
        if (glyphId <= 0) {
            throw new ArgumentOutOfRangeException(nameof(glyphId), "Shaped PDF glyph identifiers must be positive.");
        }

        GlyphId = glyphId;
        Guard.NotNull(unicodeText, nameof(unicodeText));
        UnicodeText = unicodeText;
        TextIndex = textIndex;
        AdvanceWidth = advanceWidth;
        OffsetX = offsetX;
        OffsetY = offsetY;
    }

    /// <summary>Font glyph identifier to write into the PDF content stream.</summary>
    public int GlyphId { get; }

    /// <summary>Original Unicode text represented by the shaped glyph for ToUnicode extraction.</summary>
    public string UnicodeText { get; }

    /// <summary>UTF-16 index in the original text where this glyph's source text begins.</summary>
    public int TextIndex { get; }

    /// <summary>Optional shaped horizontal advance in font design units; null uses the font's nominal glyph width.</summary>
    public int? AdvanceWidth { get; }

    /// <summary>Horizontal placement offset in font design units.</summary>
    public int OffsetX { get; }

    /// <summary>Vertical placement offset in font design units.</summary>
    public int OffsetY { get; }
}
