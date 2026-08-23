using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Describes one embedded font candidate used when planning generated PDF text fallback.
/// </summary>
public sealed class PdfEmbeddedFontFallbackCandidate {
    private readonly PdfEmbeddedFontFamily _fontFamily;

    /// <summary>
    /// Creates a fallback candidate from TrueType or OpenType/CFF font bytes.
    /// </summary>
    /// <param name="fontName">Display name used in fallback segments and diagnostics.</param>
    /// <param name="trueTypeFont">TrueType or OpenType/CFF font bytes to inspect for Unicode glyph coverage.</param>
    public PdfEmbeddedFontFallbackCandidate(string fontName, byte[] trueTypeFont)
        : this(fontName, trueTypeFont, OfficeFontUnicodeRangeSet.All) {
    }

    /// <summary>
    /// Creates a fallback candidate whose glyph coverage is limited to an explicit Unicode range policy.
    /// </summary>
    /// <param name="fontName">Display name used in fallback segments and diagnostics.</param>
    /// <param name="trueTypeFont">TrueType or OpenType/CFF font bytes to inspect for Unicode glyph coverage.</param>
    /// <param name="unicodeRanges">Unicode scalars this candidate is allowed to serve.</param>
    public PdfEmbeddedFontFallbackCandidate(
        string fontName,
        byte[] trueTypeFont,
        OfficeFontUnicodeRangeSet unicodeRanges)
        : this(fontName, trueTypeFont, unicodeRanges, OfficeFontStyle.Regular) {
    }

    internal PdfEmbeddedFontFallbackCandidate(
        string fontName,
        byte[] trueTypeFont,
        OfficeFontUnicodeRangeSet unicodeRanges,
        OfficeFontStyle style,
        string? plannerFamilyName = null) {
        Guard.NotNullOrWhiteSpace(fontName, nameof(fontName));
        Guard.NotNull(trueTypeFont, nameof(trueTypeFont));
        Guard.NotNull(unicodeRanges, nameof(unicodeRanges));
        if (trueTypeFont.Length == 0) {
            throw new ArgumentException("Embedded font fallback data cannot be empty.", nameof(trueTypeFont));
        }

        FontName = fontName;
        PlannerFamilyName = string.IsNullOrWhiteSpace(plannerFamilyName)
            ? fontName
            : plannerFamilyName!.Trim();
        _fontFamily = new PdfEmbeddedFontFamily(FontName, trueTypeFont);
        UnicodeRanges = unicodeRanges;
        Style = style & (OfficeFontStyle.Bold | OfficeFontStyle.Italic);
    }

    /// <summary>Display name used in fallback segments and diagnostics.</summary>
    public string FontName { get; }

    /// <summary>Unicode scalars this candidate is allowed to serve.</summary>
    public OfficeFontUnicodeRangeSet UnicodeRanges { get; }

    internal OfficeFontStyle Style { get; }

    internal string PlannerFamilyName { get; }

    internal byte[] DataSnapshot => _fontFamily.RegularSnapshot;

    internal PdfEmbeddedFontFamily FontFamilySnapshot => _fontFamily;
}
