using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>
/// Decoded font program used by OfficeIMO measurement and outline rendering.
/// Implementations may be supplied by optional packages for formats that are not decoded by the
/// dependency-free core, including WOFF 2, CFF/CFF2, and variable-font instances.
/// </summary>
public interface IOfficeFontProgram {
    /// <summary>
    /// Stable provider-defined identity for the decoded program and selected instance. The value
    /// must change when rendering-relevant configuration, such as variable-font axes, changes.
    /// </summary>
    string Fingerprint { get; }

    /// <summary>Best available human-readable face name.</summary>
    string? DisplayName { get; }

    /// <summary>Selected collection face index, when the source is a font collection.</summary>
    int? CollectionIndex { get; }

    /// <summary>Font design units per em.</summary>
    int UnitsPerEm { get; }

    /// <summary>Whether the source uses CFF or CFF2 outlines.</summary>
    bool IsOpenTypeCff { get; }

    /// <summary>
    /// Whether measurement and contours already apply complete script shaping and bidirectional layout.
    /// </summary>
    bool ProvidesComplexTextLayout { get; }

    /// <summary>Returns an independent font-data snapshot suitable for an external shaping provider.</summary>
    byte[] GetFontDataForShaping();

    /// <summary>Tests whether the selected program contains glyphs for every text element.</summary>
    bool HasGlyphs(string text);

    /// <summary>Measures text at the requested font size.</summary>
    double Measure(string text, double fontSize);

    /// <summary>Measures the supplied grapheme-safe text elements.</summary>
    IReadOnlyList<double> MeasureTextElements(IReadOnlyList<string> elements, double fontSize);

    /// <summary>Returns the line height at the requested font size.</summary>
    double LineHeight(double fontSize);

    /// <summary>Returns the face line-spacing ratio.</summary>
    double LineSpacingRatio { get; }

    /// <summary>Returns filled contours for the supplied text.</summary>
    List<List<OfficePoint>> GetTextContours(string text, double x, double y, double fontSize);

    /// <summary>Maps one Unicode scalar to a glyph and design-unit advance.</summary>
    bool TryGetGlyphMetrics(int scalar, out int glyphId, out int advanceWidth);

    /// <summary>Measures a validated externally shaped run.</summary>
    double MeasureShapedText(string text, OfficeTextShapingResult result, double fontSize);

    /// <summary>Returns filled contours for a validated externally shaped run.</summary>
    List<List<OfficePoint>> GetShapedTextContours(
        string text,
        OfficeTextShapingResult result,
        double x,
        double y,
        double fontSize);
}

/// <summary>
/// Optional bounded outline contract for font programs that can stop contour expansion promptly.
/// Renderers use this seam for cancellation and untrusted-input output budgets while retaining
/// compatibility with existing <see cref="IOfficeFontProgram"/> implementations.
/// </summary>
public interface IOfficeBoundedFontProgram : IOfficeFontProgram {
    /// <summary>Returns contours while enforcing the maximum total point count.</summary>
    List<List<OfficePoint>> GetTextContoursBounded(
        string text,
        double x,
        double y,
        double fontSize,
        int maximumPointCount,
        CancellationToken cancellationToken);

    /// <summary>Returns externally shaped contours while enforcing the maximum total point count.</summary>
    List<List<OfficePoint>> GetShapedTextContoursBounded(
        string text,
        OfficeTextShapingResult result,
        double x,
        double y,
        double fontSize,
        int maximumPointCount,
        CancellationToken cancellationToken);
}

/// <summary>Request passed to an optional font-program provider.</summary>
public sealed class OfficeFontProgramLoadRequest {
    private readonly byte[] _data;

    internal OfficeFontProgramLoadRequest(
        string familyName,
        byte[] data,
        OfficeFontStyle style,
        OfficeFontContainerFormat containerFormat,
        int maximumDecodedBytes) {
        FamilyName = familyName;
        _data = (byte[])data.Clone();
        Style = style;
        ContainerFormat = containerFormat;
        MaximumDecodedBytes = maximumDecodedBytes;
    }

    /// <summary>CSS/Office family requested by the caller.</summary>
    public string FamilyName { get; }

    /// <summary>Requested face style.</summary>
    public OfficeFontStyle Style { get; }

    /// <summary>Detected source container.</summary>
    public OfficeFontContainerFormat ContainerFormat { get; }

    /// <summary>Maximum decoded bytes the provider may retain for this face.</summary>
    public int MaximumDecodedBytes { get; }

    /// <summary>Returns an independent copy of the bounded source bytes.</summary>
    public byte[] Data => (byte[])_data.Clone();

    internal byte[] DataSnapshot => _data;
}

/// <summary>Successful result returned by an optional font-program provider.</summary>
public sealed class OfficeFontProgramLoadResult {
    private readonly byte[]? _staticOpenTypeData;

    /// <summary>Creates a successful provider result.</summary>
    /// <param name="program">Decoded measurement and outline program.</param>
    /// <param name="decodedByteCount">Total decoded bytes retained by the program.</param>
    /// <param name="staticOpenTypeData">
    /// Optional independent sfnt snapshot that is safe to embed as a non-variable PDF font.
    /// Leave null for WOFF 2-only, CFF2, or active variable-font instances that must be outlined.
    /// </param>
    public OfficeFontProgramLoadResult(
        IOfficeFontProgram program,
        int decodedByteCount,
        byte[]? staticOpenTypeData = null) {
        Program = program ?? throw new ArgumentNullException(nameof(program));
        if (decodedByteCount <= 0) throw new ArgumentOutOfRangeException(nameof(decodedByteCount));
        DecodedByteCount = decodedByteCount;
        _staticOpenTypeData = staticOpenTypeData == null ? null : (byte[])staticOpenTypeData.Clone();
    }

    /// <summary>Decoded measurement and outline program.</summary>
    public IOfficeFontProgram Program { get; }

    /// <summary>Total decoded bytes retained by the program.</summary>
    public int DecodedByteCount { get; }

    /// <summary>Whether the provider supplied a static sfnt program safe for PDF embedding.</summary>
    public bool CanEmbedAsStaticPdfFont => _staticOpenTypeData != null;

    /// <summary>Returns an independent static sfnt snapshot when PDF embedding is safe.</summary>
    public byte[]? StaticOpenTypeData =>
        _staticOpenTypeData == null ? null : (byte[])_staticOpenTypeData.Clone();

    internal byte[]? StaticOpenTypeDataSnapshot => _staticOpenTypeData;
}

/// <summary>
/// Optional decoder and outline-engine seam for font formats outside the dependency-free core.
/// </summary>
public interface IOfficeFontProgramProvider {
    /// <summary>
    /// Attempts to load a bounded font program. Return <see langword="null"/> when the provider does
    /// not support the supplied container or face.
    /// </summary>
    OfficeFontProgramLoadResult? TryLoad(OfficeFontProgramLoadRequest request);
}
