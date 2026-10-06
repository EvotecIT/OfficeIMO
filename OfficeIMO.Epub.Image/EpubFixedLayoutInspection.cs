using OfficeIMO.Drawing;
using OfficeIMO.Html;

namespace OfficeIMO.Epub.Image;

/// <summary>Managed layout evidence for one fixed-layout XHTML chapter, not reading-system certification.</summary>
public sealed class EpubFixedLayoutInspection {
    internal EpubFixedLayoutInspection(string path, double width, double height, HtmlRenderDocument rendering,
        OfficeDrawingQualityReport quality, IEnumerable<EpubDiagnostic> packageDiagnostics,
        IReadOnlyList<OfficeImageExportDiagnostic> preparationDiagnostics, EpubFixedLayoutRegionInspection[] regions) {
        Path = path; ViewportWidth = width; ViewportHeight = height; Rendering = rendering; CanvasQuality = quality;
        PackageDiagnostics = Array.AsReadOnly(packageDiagnostics.ToArray());
        PreparationDiagnostics = preparationDiagnostics;
        Regions = Array.AsReadOnly(regions);
    }
    /// <summary>Package-relative chapter path.</summary>
    public string Path { get; }
    /// <summary>Declared viewport width in CSS pixels.</summary>
    public double ViewportWidth { get; }
    /// <summary>Declared viewport height in CSS pixels.</summary>
    public double ViewportHeight { get; }
    /// <summary>Rendered scene, fonts and all renderer diagnostics. Its continuous surface can exceed the viewport.</summary>
    public HtmlRenderDocument Rendering { get; }
    /// <summary>Rendered element rectangle findings against the declared canvas. This is not glyph-ink,
    /// shadow/filter, clipped-content or individual-region overflow measurement.</summary>
    public OfficeDrawingQualityReport CanvasQuality { get; }
    /// <summary>Requested region inspections in caller order. Empty when only the page canvas was inspected.</summary>
    public IReadOnlyList<EpubFixedLayoutRegionInspection> Regions { get; }
    /// <summary>Package and extraction diagnostics retained regardless of image-export suppression options.</summary>
    public IReadOnlyList<EpubDiagnostic> PackageDiagnostics { get; }
    /// <summary>Source preparation diagnostics, including unavailable source content.</summary>
    public IReadOnlyList<OfficeImageExportDiagnostic> PreparationDiagnostics { get; }
    /// <summary>Whether a rendered element rectangle extends outside the declared viewport.</summary>
    public bool HasCanvasOverflow => CanvasQuality.Issues.Any(issue => issue.Kind == OfficeDrawingQualityIssueKind.ElementOutsideBounds);
    /// <summary>Whether package, preparation or rendering reports a warning/error or diagnosed rendering loss.
    /// False does not establish native-reader or pixel-level fidelity.</summary>
    public bool HasRenderingWarnings => Rendering.HasLoss || PackageDiagnostics.Any(d => d.Severity != EpubDiagnosticSeverity.Info) ||
        PreparationDiagnostics.Any(d => d.Severity != OfficeImageExportDiagnosticSeverity.Info);
}
