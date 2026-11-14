namespace OfficeIMO;

/// <summary>One stable, machine-readable accessibility finding in a Word, PowerPoint or Excel document.</summary>
public sealed class OfficeAccessibilityFinding {
    internal OfficeAccessibilityFinding(OfficeAccessibilitySeverity severity, string code, string message, string location,
        int? locationIndex = null, uint? objectId = null, string? objectName = null, bool isHeuristic = false,
        double? measuredValue = null, double? requiredValue = null) {
        Severity = severity;
        Code = string.IsNullOrWhiteSpace(code) ? "accessibility-unknown" : code;
        Message = message ?? string.Empty;
        Location = string.IsNullOrWhiteSpace(location) ? "Document" : location;
        LocationIndex = locationIndex;
        ObjectId = objectId;
        ObjectName = objectName;
        IsHeuristic = isHeuristic;
        MeasuredValue = measuredValue;
        RequiredValue = requiredValue;
    }

    /// <summary>Finding severity.</summary>
    public OfficeAccessibilitySeverity Severity { get; }
    /// <summary>Stable diagnostic code suitable for CI policies. Word, PowerPoint and Excel use the same code for the same rule.</summary>
    public string Code { get; }
    /// <summary>Human-readable explanation.</summary>
    public string Message { get; }
    /// <summary>Where the finding is, such as <c>Document</c>, <c>Paragraph 12</c>, <c>Table 2</c>, <c>Slide 3</c> or <c>Data!B3</c>.</summary>
    public string Location { get; }
    /// <summary>Zero-based index of the paragraph (Word), slide (PowerPoint) or worksheet (Excel) the finding belongs to, when it has one.</summary>
    public int? LocationIndex { get; }
    /// <summary>Drawing object identifier of the affected image, shape, chart or table (for Excel tables, the table id), when applicable.</summary>
    public uint? ObjectId { get; }
    /// <summary>Name of the affected image, shape, chart or table, when applicable.</summary>
    public string? ObjectName { get; }
    /// <summary>Whether the finding comes from a heuristic, such as guessing that a table is used for layout.</summary>
    public bool IsHeuristic { get; }
    /// <summary>Measured value, such as a contrast ratio.</summary>
    public double? MeasuredValue { get; }
    /// <summary>Value the policy requires.</summary>
    public double? RequiredValue { get; }

    /// <inheritdoc />
    public override string ToString() => Severity + " " + Code + " at " + Location + ": " + Message;
}
