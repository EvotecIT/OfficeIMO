namespace OfficeIMO.Tool.Agent;

/// <summary>Bounded PDF automation outcome. Contains artifact metadata and diagnostic codes, never extracted or recognized text.</summary>
public sealed class AgentPdfWorkflowResult {
    public string Operation { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
    public string FailureKind { get; set; } = string.Empty;
    public bool Succeeded { get; set; }
    public string? Summary { get; set; }
    public string? OutputPath { get; set; }
    public long OutputBytes { get; set; }
    public int ArtifactCount { get; set; }
    public int AddedWordCount { get; set; }
    public int DiagnosticCount { get; set; }
    public IReadOnlyList<AgentPdfArtifact> Artifacts { get; set; } = Array.Empty<AgentPdfArtifact>();
    public IReadOnlyList<AgentPdfDiagnostic> Diagnostics { get; set; } = Array.Empty<AgentPdfDiagnostic>();
    public bool Truncated { get; set; }
}

/// <summary>One generated artifact; a bounded response may omit additional artifacts from the same output folder.</summary>
public sealed class AgentPdfArtifact {
    public string Path { get; set; } = string.Empty;
    public long SizeBytes { get; set; }
    public int? FirstSourcePage { get; set; }
    public int? PageCount { get; set; }
}

/// <summary>Diagnostic identity and severity without provider messages or document content.</summary>
public sealed class AgentPdfDiagnostic {
    public string Code { get; set; } = string.Empty;
    public string Severity { get; set; } = string.Empty;
}

/// <summary>Bounded discovery of OCR providers trusted by this server process.</summary>
public sealed class AgentPdfOcrProvidersResult {
    public int ProviderCount { get; set; }
    public IReadOnlyList<string> ProviderIds { get; set; } = Array.Empty<string>();
    public bool Truncated { get; set; }
}
