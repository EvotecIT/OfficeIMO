using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// One caller-supplied TrueType face that can be reused by dependency-free drawing renderers.
/// </summary>
public sealed class OfficeFontFace {
    private readonly byte[] _data;

    internal OfficeFontFace(string familyName, byte[] data, OfficeFontStyle style, OfficeTrueTypeFont parsedFont) {
        FamilyName = familyName;
        Style = NormalizeStyle(style);
        _data = (byte[])data.Clone();
        ParsedFont = parsedFont;
    }

    /// <summary>CSS/Office family name used to select the face.</summary>
    public string FamilyName { get; }

    /// <summary>Bold and italic face attributes.</summary>
    public OfficeFontStyle Style { get; }

    /// <summary>Independent copy of the original TrueType bytes.</summary>
    public byte[] Data => (byte[])_data.Clone();

    internal byte[] DataSnapshot => _data;

    internal OfficeTrueTypeFont ParsedFont { get; }

    internal OfficeFontFace Clone() => new OfficeFontFace(FamilyName, _data, Style, ParsedFont);

    internal static OfficeFontStyle NormalizeStyle(OfficeFontStyle style) =>
        style & (OfficeFontStyle.Bold | OfficeFontStyle.Italic);
}
