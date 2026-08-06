using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// One caller-supplied TrueType face that can be reused by dependency-free drawing renderers.
/// </summary>
public sealed class OfficeFontFace {
    private readonly byte[] _data;

    internal OfficeFontFace(
        string familyName,
        string resourceFamilyName,
        byte[] data,
        OfficeFontStyle style,
        OfficeFontUnicodeRangeSet unicodeRanges,
        OfficeTrueTypeFont parsedFont,
        bool useDataSnapshot = false) {
        FamilyName = familyName;
        ResourceFamilyName = resourceFamilyName;
        Style = NormalizeStyle(style);
        UnicodeRanges = unicodeRanges;
        _data = useDataSnapshot ? data : (byte[])data.Clone();
        ParsedFont = parsedFont;
    }

    /// <summary>CSS/Office family name used to select the face.</summary>
    public string FamilyName { get; }

    /// <summary>Unique family name used by exporters after unicode-range selection.</summary>
    public string ResourceFamilyName { get; }

    /// <summary>Bold and italic face attributes.</summary>
    public OfficeFontStyle Style { get; }

    /// <summary>Unicode scalars this face is allowed to serve.</summary>
    public OfficeFontUnicodeRangeSet UnicodeRanges { get; }

    /// <summary>Independent copy of normalized OpenType bytes.</summary>
    public byte[] Data => (byte[])_data.Clone();

    internal byte[] DataSnapshot => _data;

    internal OfficeTrueTypeFont ParsedFont { get; }

    internal OfficeFontFace Clone() =>
        new OfficeFontFace(FamilyName, ResourceFamilyName, _data, Style, UnicodeRanges, ParsedFont, useDataSnapshot: true);

    internal OfficeFontFace CreateAlias(string familyName, string resourceFamilyName) =>
        new OfficeFontFace(familyName, resourceFamilyName, _data, Style, UnicodeRanges, ParsedFont, useDataSnapshot: true);

    internal bool Covers(string text) => UnicodeRanges.ContainsText(text) && ParsedFont.HasGlyphs(text);

    internal bool HasGlyphs(string text) => ParsedFont.HasGlyphs(text);

    internal static OfficeFontStyle NormalizeStyle(OfficeFontStyle style) =>
        style & (OfficeFontStyle.Bold | OfficeFontStyle.Italic);
}
