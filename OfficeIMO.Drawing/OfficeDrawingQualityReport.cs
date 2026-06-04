using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Drawing;

/// <summary>
/// Result of shared drawing quality analysis.
/// </summary>
public sealed class OfficeDrawingQualityReport {
    internal OfficeDrawingQualityReport(IReadOnlyList<OfficeDrawingQualityIssue> issues) {
        Issues = new ReadOnlyCollection<OfficeDrawingQualityIssue>(new List<OfficeDrawingQualityIssue>(issues));
    }

    /// <summary>Detected quality issues.</summary>
    public IReadOnlyList<OfficeDrawingQualityIssue> Issues { get; }

    /// <summary>Whether the drawing has any detected quality issues.</summary>
    public bool HasIssues => Issues.Count > 0;
}
