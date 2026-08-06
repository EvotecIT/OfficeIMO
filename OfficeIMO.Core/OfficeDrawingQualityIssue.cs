using System.Globalization;

namespace OfficeIMO.Drawing;

/// <summary>
/// Structured drawing quality diagnostic emitted by <see cref="OfficeDrawingQualityAnalyzer"/>.
/// </summary>
public sealed class OfficeDrawingQualityIssue {
    /// <summary>
    /// Creates a drawing quality issue.
    /// </summary>
    public OfficeDrawingQualityIssue(OfficeDrawingQualityIssueKind kind, string message, int elementIndex, int? relatedElementIndex = null) {
        Kind = kind;
        Message = message ?? string.Empty;
        ElementIndex = elementIndex;
        RelatedElementIndex = relatedElementIndex;
    }

    /// <summary>Issue kind.</summary>
    public OfficeDrawingQualityIssueKind Kind { get; }

    /// <summary>Human-readable diagnostic message.</summary>
    public string Message { get; }

    /// <summary>Zero-based drawing element index.</summary>
    public int ElementIndex { get; }

    /// <summary>Optional related zero-based drawing element index.</summary>
    public int? RelatedElementIndex { get; }

    /// <inheritdoc />
    public override string ToString() {
        return RelatedElementIndex.HasValue
            ? string.Format(CultureInfo.InvariantCulture, "{0} at element {1} related to element {2}: {3}", Kind, ElementIndex, RelatedElementIndex.Value, Message)
            : string.Format(CultureInfo.InvariantCulture, "{0} at element {1}: {2}", Kind, ElementIndex, Message);
    }
}
