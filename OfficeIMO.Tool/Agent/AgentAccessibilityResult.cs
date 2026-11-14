namespace OfficeIMO.Tool.Agent;

/// <summary>Accessibility inspection of one Word, Excel or PowerPoint file, using the report shared by all three formats.</summary>
public sealed class AgentAccessibilityResult {
    /// <summary>Messages, locations and object names can quote document text, such as link labels and sheet names.</summary>
    public string ContentTrust => "untrusted";
    public string Path { get; set; } = string.Empty;
    public string Format { get; set; } = string.Empty;
    public string Profile { get; set; } = string.Empty;
    public int SchemaVersion { get; set; }
    public int ErrorCount { get; set; }
    public int WarningCount { get; set; }
    public bool IsSuccessful { get; set; }
    public int FindingCount { get; set; }
    public IReadOnlyList<AgentAccessibilityFinding> Findings { get; set; } = Array.Empty<AgentAccessibilityFinding>();
    public bool Truncated { get; set; }
    /// <summary>SARIF 2.1.0 log for every finding, when requested.</summary>
    public string? Sarif { get; set; }
}

/// <summary>One accessibility finding with its stable rule code.</summary>
public sealed class AgentAccessibilityFinding {
    public string Severity { get; set; } = string.Empty;
    public string Code { get; set; } = string.Empty;
    public string Message { get; set; } = string.Empty;
    public string Location { get; set; } = string.Empty;
    public int? LocationIndex { get; set; }
    public uint? ObjectId { get; set; }
    public string? ObjectName { get; set; }
    public bool IsHeuristic { get; set; }
    public double? MeasuredValue { get; set; }
    public double? RequiredValue { get; set; }
}
