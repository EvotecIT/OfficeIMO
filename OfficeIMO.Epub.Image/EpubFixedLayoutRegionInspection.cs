using OfficeIMO.Drawing;

namespace OfficeIMO.Epub.Image;

/// <summary>Rendered element rectangle findings in one region's local border-box coordinate space.
/// This does not measure glyph ink, shadows, filters or hidden clipped content.</summary>
public sealed class EpubFixedLayoutRegionInspection {
    internal EpubFixedLayoutRegionInspection(string elementId, OfficeDrawingQualityReport quality) {
        ElementId = elementId; Quality = quality;
    }
    /// <summary>Exact authored HTML element ID.</summary>
    public string ElementId { get; }
    /// <summary>Bounds findings after descendant transforms and authored clips, before region/ancestor transforms.</summary>
    public OfficeDrawingQualityReport Quality { get; }
    /// <summary>Whether a rendered element rectangle extends outside the region's local border box.</summary>
    public bool HasOverflow => Quality.Issues.Any(issue => issue.Kind == OfficeDrawingQualityIssueKind.ElementOutsideBounds);
}
