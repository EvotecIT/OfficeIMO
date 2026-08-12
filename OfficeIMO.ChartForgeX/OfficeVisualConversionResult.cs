using System;
using System.Collections.Generic;
using global::ChartForgeX.VisualArtifacts;
using OfficeIMO.Drawing;

namespace OfficeIMO.ChartForgeX;

/// <summary>Contains reusable Office representations and metadata for one ChartForgeX visual artifact.</summary>
public sealed class OfficeVisualConversionResult {
    private readonly byte[] _svgBytes;
    private readonly byte[] _placementBytes;

    internal OfficeVisualConversionResult(
        VisualArtifact? artifact,
        string id,
        string title,
        byte[] svgBytes,
        byte[] placementBytes,
        OfficeVisualMediaFormat placementFormat,
        OfficeDrawing drawing,
        double widthPoints,
        double heightPoints,
        string alternativeText,
        bool isDecorative,
        OfficeVisualSvgPolicy svgPolicy,
        IReadOnlyList<OfficeVisualRegion> regions,
        OfficeVisualConversionReport report) {
        Artifact = artifact;
        Id = id;
        Title = title;
        _svgBytes = svgBytes;
        _placementBytes = placementBytes;
        PlacementFormat = placementFormat;
        Drawing = drawing;
        WidthPoints = widthPoints;
        HeightPoints = heightPoints;
        AlternativeText = alternativeText;
        IsDecorative = isDecorative;
        SvgPolicy = svgPolicy;
        Regions = regions;
        Report = report;
    }

    /// <summary>Gets the source visual artifact.</summary>
    public VisualArtifact? Artifact { get; }

    /// <summary>Gets the stable visual identifier.</summary>
    public string Id { get; }

    /// <summary>Gets the visual title when one was supplied.</summary>
    public string Title { get; }

    /// <summary>Gets the Office drawing representation used by PDF and drawing consumers.</summary>
    public OfficeDrawing Drawing { get; }

    /// <summary>Gets the format of the payload used by image-based Office placements.</summary>
    public OfficeVisualMediaFormat PlacementFormat { get; }

    /// <summary>Gets the MIME type of the payload used by image-based Office placements.</summary>
    public string PlacementMediaType => PlacementFormat == OfficeVisualMediaFormat.Svg ? "image/svg+xml" : "image/png";

    /// <summary>Gets the file extension of the payload used by image-based Office placements.</summary>
    public string PlacementFileExtension => PlacementFormat == OfficeVisualMediaFormat.Svg ? ".svg" : ".png";

    /// <summary>Gets the resolved output width in points.</summary>
    public double WidthPoints { get; }

    /// <summary>Gets the resolved output height in points.</summary>
    public double HeightPoints { get; }

    /// <summary>Gets the resolved accessible description used by document placement helpers.</summary>
    public string AlternativeText { get; }

    /// <summary>Gets whether the source visual is decorative and should be omitted from accessibility structure.</summary>
    public bool IsDecorative { get; }

    /// <summary>Gets the fidelity policy selected for this conversion.</summary>
    public OfficeVisualSvgPolicy SvgPolicy { get; }

    /// <summary>Gets artifact regions transformed to Office point coordinates.</summary>
    public IReadOnlyList<OfficeVisualRegion> Regions { get; }

    /// <summary>Gets the conversion fidelity report.</summary>
    public OfficeVisualConversionReport Report { get; }

    /// <summary>Returns an independent copy of the rendered SVG payload.</summary>
    public byte[] GetSvgBytes() => (byte[])_svgBytes.Clone();

    /// <summary>Returns an independent copy of the payload selected for Word, Excel, and PowerPoint placement.</summary>
    public byte[] GetPlacementBytes() => (byte[])_placementBytes.Clone();
}
