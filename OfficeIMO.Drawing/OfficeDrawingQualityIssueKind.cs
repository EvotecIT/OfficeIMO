namespace OfficeIMO.Drawing;

/// <summary>
/// Kind of drawing quality issue detected by <see cref="OfficeDrawingQualityAnalyzer"/>.
/// </summary>
public enum OfficeDrawingQualityIssueKind {
    /// <summary>An element extends outside the drawing canvas bounds.</summary>
    ElementOutsideBounds,

    /// <summary>Two text boxes overlap in a way that can lead to unreadable output.</summary>
    TextOverlap
}
