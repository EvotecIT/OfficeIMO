using OfficeIMO.Drawing;
using OfficeIMO.Html;

namespace OfficeIMO.Epub.Image;

/// <summary>Rendered rectangle and positioned XHTML text-ink findings in one region's local border-box space.
/// This does not measure shadows, filters or hidden clipped content.</summary>
public sealed class EpubFixedLayoutRegionInspection {
    internal EpubFixedLayoutRegionInspection(string elementId, OfficeDrawingQualityReport quality, IReadOnlyList<HtmlDiagnostic> textInkDiagnostics) {
        ElementId = elementId; Quality = quality; TextInkDiagnostics = textInkDiagnostics;
    }
    /// <summary>Exact authored HTML element ID.</summary>
    public string ElementId { get; }
    /// <summary>Bounds findings after descendant transforms and authored clips, before region/ancestor transforms.</summary>
    public OfficeDrawingQualityReport Quality { get; }
    /// <summary>Positioned XHTML text paint bounds at nominal CSS-pixel scale 1, in local coordinates.
    /// Includes non-zero glyph winding, descendant transforms, rectangular/convex contour clipping and conservative decoration bounds. Unsupported path clips, vector text and
    /// unavailable outlines are diagnosed as unmeasured; ancestor transforms and clips are excluded.</summary>
    public IReadOnlyList<HtmlDiagnostic> TextInkDiagnostics { get; }
    /// <summary>Whether measured text paint bounds extend outside this region's local border box.</summary>
    public bool HasTextInkOverflow => TextInkDiagnostics.Any(d => d.Code == HtmlRenderDiagnosticCodes.TextInkOutsideRegion);
    /// <summary>Whether a rendered element rectangle extends outside the region's local border box.</summary>
    public bool HasOverflow => Quality.Issues.Any(issue => issue.Kind == OfficeDrawingQualityIssueKind.ElementOutsideBounds);
}
