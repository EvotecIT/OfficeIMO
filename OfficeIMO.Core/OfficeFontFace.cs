using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>
/// One caller-supplied font face that can be reused by drawing renderers.
/// </summary>
public sealed partial class OfficeFontFace {
    private readonly byte[] _data;
    private readonly bool _automaticOpticalSizing;

    internal OfficeFontFace(
        string familyName,
        string resourceFamilyName,
        byte[] data,
        OfficeFontStyle style,
        OfficeFontFaceDescriptor descriptor,
        OfficeFontUnicodeRangeSet unicodeRanges,
        IOfficeFontProgram parsedFont,
        OfficeFontContainerFormat containerFormat,
        bool canEmbedAsStaticPdfFont,
        bool useDataSnapshot = false,
        bool automaticOpticalSizing = true) {
        FamilyName = familyName;
        ResourceFamilyName = resourceFamilyName;
        Descriptor = descriptor;
        Style = NormalizeStyle(style);
        UnicodeRanges = unicodeRanges;
        _data = useDataSnapshot ? data : (byte[])data.Clone();
        ParsedFont = parsedFont;
        ContainerFormat = containerFormat;
        CanEmbedAsStaticPdfFont = canEmbedAsStaticPdfFont;
        _automaticOpticalSizing = automaticOpticalSizing;
    }

    /// <summary>CSS/Office family name used to select the face.</summary>
    public string FamilyName { get; }

    /// <summary>Unique family name used by exporters after unicode-range selection.</summary>
    public string ResourceFamilyName { get; }

    /// <summary>Bold and italic face attributes.</summary>
    public OfficeFontStyle Style { get; }

    /// <summary>Numeric weight, stretch, and slant attributes used for CSS-compatible face selection.</summary>
    public OfficeFontFaceDescriptor Descriptor { get; }

    /// <summary>Unicode scalars this face is allowed to serve.</summary>
    public OfficeFontUnicodeRangeSet UnicodeRanges { get; }

    /// <summary>
    /// Independent copy of the accepted face bytes. This is normalized sfnt data for built-in
    /// TrueType/WOFF 1 faces and provider-selected bytes for other formats.
    /// </summary>
    public byte[] Data => (byte[])_data.Clone();

    /// <summary>Detected source container format.</summary>
    public OfficeFontContainerFormat ContainerFormat { get; }

    /// <summary>Whether <see cref="Data"/> is a static sfnt program safe for direct PDF embedding.</summary>
    public bool CanEmbedAsStaticPdfFont { get; }

    /// <summary>
    /// Decoded measurement, shaping, and outline program for this face. Exporters can use this
    /// when direct static-font embedding is unavailable, for example for WOFF 2, CFF2, or an
    /// active variable-font instance.
    /// </summary>
    public IOfficeFontProgram Program => ParsedFont;

    internal byte[] DataSnapshot => _data;

    internal IOfficeFontProgram ParsedFont { get; }

    // Library-selected installed substitutes retain their origin across collection
    // snapshots. An explicit caller-created alias remains a caller-supplied family.
    internal string? InstalledSubstituteFamily { get; set; }

    internal IReadOnlyDictionary<string, float>? VariationCoordinatesForShaping =>
        (ParsedFont as IOfficeVariableFontProgram)?.VariationCoordinatesForShaping;

    internal OfficeFontFace Clone() =>
        new OfficeFontFace(FamilyName, ResourceFamilyName, _data, Style, Descriptor, UnicodeRanges, ParsedFont, ContainerFormat, CanEmbedAsStaticPdfFont, useDataSnapshot: true, automaticOpticalSizing: _automaticOpticalSizing) {
            InstalledSubstituteFamily = InstalledSubstituteFamily
        };

    internal OfficeFontFace CreateAlias(string familyName, string resourceFamilyName) =>
        new OfficeFontFace(familyName, resourceFamilyName, _data, Style, Descriptor, UnicodeRanges, ParsedFont, ContainerFormat, CanEmbedAsStaticPdfFont, useDataSnapshot: true, automaticOpticalSizing: _automaticOpticalSizing);

    internal bool Covers(string text) => UnicodeRanges.ContainsFontCoverageText(text) && ParsedFont.HasGlyphs(text);

    internal bool HasGlyphs(string text) => ParsedFont.HasGlyphs(text);

    internal static OfficeFontStyle NormalizeStyle(OfficeFontStyle style) =>
        style & (OfficeFontStyle.Bold | OfficeFontStyle.Italic);
}
