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
        _regularFaceCandidate = new Lazy<PdfEmbeddedFontFallbackCandidate>(CreateRegularFaceCandidate);
        _styledFaceCoverageTests = CreateStyledFaceCoverageTests();
    }

    /// <summary>
    /// Creates a fallback candidate from a font family. Coverage is planned against the regular face, and
    /// bold, italic and bold-italic runs are painted with the family's own faces. Characters a styled face
    /// lacks, and styles the family has no face for, are synthesized from the regular face.
    /// </summary>
    /// <param name="fontFamily">Font family whose regular face decides coverage and whose styled faces paint styled runs.</param>
    public PdfEmbeddedFontFallbackCandidate(PdfEmbeddedFontFamily fontFamily)
        : this(fontFamily, null, OfficeFontStyle.Regular, null) {
    }

    internal PdfEmbeddedFontFallbackCandidate(
        PdfEmbeddedFontFamily fontFamily,
        OfficeFontUnicodeRangeSet? unicodeRanges = null,
        OfficeFontStyle style = OfficeFontStyle.Regular,
        string? plannerFamilyName = null) {
        Guard.NotNull(fontFamily, nameof(fontFamily));

        FontName = fontFamily.FamilyName;
        PlannerFamilyName = string.IsNullOrWhiteSpace(plannerFamilyName)
            ? FontName
            : plannerFamilyName!.Trim();
        _fontFamily = fontFamily.CreateFallbackSnapshot();
        UnicodeRanges = unicodeRanges ?? OfficeFontUnicodeRangeSet.All;
        Style = style & (OfficeFontStyle.Bold | OfficeFontStyle.Italic);
        _regularFaceCandidate = new Lazy<PdfEmbeddedFontFallbackCandidate>(CreateRegularFaceCandidate);
        _styledFaceCoverageTests = CreateStyledFaceCoverageTests();
    }

    /// <summary>Display name used in fallback segments and diagnostics.</summary>
    public string FontName { get; }

    /// <summary>Unicode scalars this candidate is allowed to serve.</summary>
    public OfficeFontUnicodeRangeSet UnicodeRanges { get; }

    internal OfficeFontStyle Style { get; }

    internal string PlannerFamilyName { get; }

    // A registration may need a different resource name, but a reusable plan
    // still selects the same immutable candidate in every options snapshot.
    internal object RegistrationIdentity { get; } = new object();

    private PdfEmbeddedFontFallbackCandidate(PdfEmbeddedFontFallbackCandidate source, string registeredFamilyName) {
        FontName = registeredFamilyName;
        PlannerFamilyName = source.PlannerFamilyName;
        _fontFamily = source._fontFamily.WithFallbackFamilyName(registeredFamilyName);
        UnicodeRanges = source.UnicodeRanges;
        Style = source.Style;
        RegistrationIdentity = source.RegistrationIdentity;
        _regularFaceCandidate = source._regularFaceCandidate;
        _styledFaceCoverageTests = source._styledFaceCoverageTests;
        IsRegularFaceCandidate = source.IsRegularFaceCandidate;
    }

    internal PdfEmbeddedFontFallbackCandidate WithRegisteredFamilyName(string familyName) =>
        new(this, familyName);

    internal byte[] DataSnapshot => _fontFamily.RegularSnapshot;

    /// <summary>
    /// The face that paints a style in a compatibility slot, chosen like a caller-registered family:
    /// a style is synthesized only when the candidate has no face for it.
    /// </summary>
    internal (byte[] Data, bool SyntheticBold, bool SyntheticOblique) GetSlotFace(bool bold, bool italic) {
        byte[] regular = _fontFamily.RegularSnapshot;
        byte[]? boldFace = _fontFamily.BoldSnapshot;
        byte[]? italicFace = _fontFamily.ItalicSnapshot;
        if (bold && italic) {
            byte[]? boldItalicFace = _fontFamily.BoldItalicSnapshot;
            return boldItalicFace != null
                ? (boldItalicFace, false, false)
                : (boldFace ?? italicFace ?? regular, boldFace == null, boldFace != null || italicFace == null);
        }
        if (bold) return (boldFace ?? regular, boldFace == null, false);
        if (italic) return (italicFace ?? regular, false, italicFace == null);
        return (regular, false, false);
    }

    private readonly Lazy<PdfEmbeddedFontFallbackCandidate> _regularFaceCandidate;

    /// <summary>True for the regular-face companion that paints styled text its own styled faces cannot.</summary>
    internal bool IsRegularFaceCandidate { get; private set; }

    /// <summary>
    /// A companion holding only the regular face. Styled text whose glyphs are missing from the matching
    /// styled face is painted with it, so the style is synthesized for those glyphs alone.
    /// </summary>
    internal PdfEmbeddedFontFallbackCandidate RegularFaceCandidate =>
        IsRegularFaceCandidate ? this : _regularFaceCandidate.Value;

    private PdfEmbeddedFontFallbackCandidate CreateRegularFaceCandidate() {
        var candidate = new PdfEmbeddedFontFallbackCandidate(
            PdfEmbeddedFontFamily.FromOwnedFaces(FontName + " Regular Face", _fontFamily.RegularSnapshot, null, null, null),
            UnicodeRanges, OfficeFontStyle.Regular, PlannerFamilyName);
        candidate.IsRegularFaceCandidate = true;
        return candidate;
    }


    /// <summary>
    /// Splits styled text into pieces the styled face paints and pieces only the regular face can paint, or
    /// returns null when the styled face paints all of it. Text the regular face covers only through shaping
    /// stays whole on the regular face so its shaped units are kept.
    /// </summary>
    internal IReadOnlyList<(string Text, bool RegularFace)>? SplitByStyledFaceCoverage(string text, bool bold, bool italic) {
        if (IsRegularFaceCandidate || (!bold && !italic) || string.IsNullOrEmpty(text)) return null;
        Func<string, StyledFaceCoverage>? coverage = _styledFaceCoverageTests[(bold ? 1 : 0) + (italic ? 2 : 0) - 1].Value;
        if (coverage == null) return null;
        StyledFaceCoverage whole = coverage(text);
        if (whole == StyledFaceCoverage.StyledFace) return null;
        if (whole == StyledFaceCoverage.ShapedRegularFace) return new[] { (text, true) };

        var pieces = new List<(string Text, bool RegularFace)>();
        var piece = new System.Text.StringBuilder();
        bool pieceRegular = false;
        foreach (string element in OfficeIMO.Drawing.OfficeTextElements.Enumerate(text)) {
            bool regular = coverage(element) != StyledFaceCoverage.StyledFace;
            if (piece.Length > 0 && regular != pieceRegular) {
                pieces.Add((piece.ToString(), pieceRegular));
                piece.Clear();
            }
            pieceRegular = regular;
            piece.Append(element);
        }
        if (piece.Length > 0) pieces.Add((piece.ToString(), pieceRegular));
        return pieces;
    }

    private readonly Lazy<Func<string, StyledFaceCoverage>?>[] _styledFaceCoverageTests;

    private Lazy<Func<string, StyledFaceCoverage>?>[] CreateStyledFaceCoverageTests() => new[] {
        new Lazy<Func<string, StyledFaceCoverage>?>(() => PdfEmbeddedFontFamily.CreateStyledFaceCoverageTest(GetSlotFace(bold: true, italic: false).Data, _fontFamily.RegularSnapshot)),
        new Lazy<Func<string, StyledFaceCoverage>?>(() => PdfEmbeddedFontFamily.CreateStyledFaceCoverageTest(GetSlotFace(bold: false, italic: true).Data, _fontFamily.RegularSnapshot)),
        new Lazy<Func<string, StyledFaceCoverage>?>(() => PdfEmbeddedFontFamily.CreateStyledFaceCoverageTest(GetSlotFace(bold: true, italic: true).Data, _fontFamily.RegularSnapshot))
    };

    internal bool HasSameFaces(PdfEmbeddedFontFallbackCandidate other) =>
        SameFace(_fontFamily.RegularSnapshot, other._fontFamily.RegularSnapshot)
        && SameFace(_fontFamily.BoldSnapshot, other._fontFamily.BoldSnapshot)
        && SameFace(_fontFamily.ItalicSnapshot, other._fontFamily.ItalicSnapshot)
        && SameFace(_fontFamily.BoldItalicSnapshot, other._fontFamily.BoldItalicSnapshot);

    private static bool SameFace(byte[]? left, byte[]? right) =>
        ReferenceEquals(left, right) || (left != null && right != null && left.SequenceEqual(right));

    internal PdfEmbeddedFontFamily FontFamilySnapshot => _fontFamily;
}
