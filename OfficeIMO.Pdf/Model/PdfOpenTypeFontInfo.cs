namespace OfficeIMO.Pdf;

/// <summary>
/// Describes a parsed OpenType or TrueType font program without claiming it can be rendered by every PDF output path.
/// </summary>
public sealed class PdfOpenTypeFontInfo {
    internal PdfOpenTypeFontInfo(
        string fontName,
        string scalerType,
        bool isOpenTypeCff,
        bool isTrueType,
        int unitsPerEm,
        int glyphCount,
        int cffTableLength,
        int unicodeScalarCount,
        IReadOnlyDictionary<int, int> unicodeCMap,
        bool hasGlyphSubstitutionTable,
        bool hasGlyphPositioningTable,
        IReadOnlyList<string> glyphSubstitutionFeatureTags,
        IReadOnlyList<string> glyphPositioningFeatureTags) {
        FontName = string.IsNullOrWhiteSpace(fontName) ? "OfficeIMOEmbeddedFont" : fontName;
        ScalerType = scalerType ?? string.Empty;
        IsOpenTypeCff = isOpenTypeCff;
        IsTrueType = isTrueType;
        UnitsPerEm = unitsPerEm;
        GlyphCount = glyphCount;
        CffTableLength = cffTableLength;
        UnicodeScalarCount = unicodeScalarCount;
        UnicodeCMap = CopyUnicodeCMap(unicodeCMap);
        HasGlyphSubstitutionTable = hasGlyphSubstitutionTable;
        HasGlyphPositioningTable = hasGlyphPositioningTable;
        GlyphSubstitutionFeatureTags = CopyFeatureTags(glyphSubstitutionFeatureTags);
        GlyphPositioningFeatureTags = CopyFeatureTags(glyphPositioningFeatureTags);
    }

    /// <summary>PostScript or configured font name sanitized for PDF font dictionaries.</summary>
    public string FontName { get; }

    /// <summary>OpenType scaler type, such as <c>OTTO</c>, <c>true</c>, or <c>0x00010000</c>.</summary>
    public string ScalerType { get; }

    /// <summary>Whether the font is an OpenType/CFF font with an <c>OTTO</c> scaler and <c>CFF </c> table.</summary>
    public bool IsOpenTypeCff { get; }

    /// <summary>Whether the font uses TrueType <c>glyf</c> outlines.</summary>
    public bool IsTrueType { get; }

    /// <summary>Units per em read from the OpenType <c>head</c> table.</summary>
    public int UnitsPerEm { get; }

    /// <summary>Number of glyphs read from the OpenType <c>maxp</c> table.</summary>
    public int GlyphCount { get; }

    /// <summary>Length of the <c>CFF </c> table in bytes, or zero for non-CFF fonts.</summary>
    public int CffTableLength { get; }

    /// <summary>Number of Unicode scalars mapped by a supported <c>cmap</c> table.</summary>
    public int UnicodeScalarCount { get; }

    /// <summary>Whether the font contains an OpenType <c>GSUB</c> glyph substitution table.</summary>
    public bool HasGlyphSubstitutionTable { get; }

    /// <summary>Whether the font contains an OpenType <c>GPOS</c> glyph positioning table.</summary>
    public bool HasGlyphPositioningTable { get; }

    /// <summary>Feature tags advertised by the font's OpenType <c>GSUB</c> feature list, such as <c>liga</c> or <c>rlig</c>.</summary>
    public IReadOnlyList<string> GlyphSubstitutionFeatureTags { get; }

    /// <summary>Feature tags advertised by the font's OpenType <c>GPOS</c> feature list, such as <c>mark</c> or <c>mkmk</c>.</summary>
    public IReadOnlyList<string> GlyphPositioningFeatureTags { get; }

    internal IReadOnlyDictionary<int, int> UnicodeCMap { get; }

    /// <summary>
    /// Checks whether a Unicode scalar maps to a non-zero glyph id in the parsed font's Unicode cmap.
    /// </summary>
    /// <param name="unicodeScalar">Unicode scalar value to check.</param>
    /// <returns>True when the scalar maps to a non-zero glyph id.</returns>
    public bool ContainsUnicodeScalar(int unicodeScalar) =>
        UnicodeCMap.TryGetValue(unicodeScalar, out int glyphId) && glyphId > 0;

    private static Dictionary<int, int> CopyUnicodeCMap(IReadOnlyDictionary<int, int> unicodeCMap) {
        Guard.NotNull(unicodeCMap, nameof(unicodeCMap));
        var copy = new Dictionary<int, int>(unicodeCMap.Count);
        foreach (KeyValuePair<int, int> entry in unicodeCMap) {
            copy[entry.Key] = entry.Value;
        }

        return copy;
    }

    private static List<string> CopyFeatureTags(IReadOnlyList<string> featureTags) {
        Guard.NotNull(featureTags, nameof(featureTags));
        var copy = new List<string>(featureTags.Count);
        foreach (string tag in featureTags) {
            if (!string.IsNullOrWhiteSpace(tag)) {
                copy.Add(tag);
            }
        }

        return copy;
    }
}
