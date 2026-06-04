using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Result of rendering a chart snapshot into shared drawing primitives.
/// </summary>
public sealed class OfficeChartRenderingResult {
    /// <summary>
    /// Creates a chart rendering result.
    /// </summary>
    public OfficeChartRenderingResult(OfficeDrawing drawing, OfficeDrawingQualityReport qualityReport) {
        Drawing = drawing ?? throw new ArgumentNullException(nameof(drawing));
        QualityReport = qualityReport ?? throw new ArgumentNullException(nameof(qualityReport));
    }

    /// <summary>Rendered chart drawing scene.</summary>
    public OfficeDrawing Drawing { get; }

    /// <summary>Reusable drawing quality diagnostics for the rendered chart scene.</summary>
    public OfficeDrawingQualityReport QualityReport { get; }
}
