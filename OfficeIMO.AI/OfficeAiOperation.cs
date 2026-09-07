namespace OfficeIMO.AI;

/// <summary>Supported read-only document operations.</summary>
public enum OfficeAiOperation {
    /// <summary>Answer a question using supplied source evidence.</summary>
    Ask,
    /// <summary>Explain selected evidence without expanding its scope.</summary>
    Explain,
    /// <summary>Summarize every selected evidence batch, reporting omissions.</summary>
    Summarize,
    /// <summary>Extract exactly the requested named fields.</summary>
    ExtractFields,
    /// <summary>Propose logical blocks and rectangular tables from image or text evidence.</summary>
    Parse
}

/// <summary>Supported normalized scalar field types.</summary>
public enum OfficeAiFieldType {
    /// <summary>Unmodified text.</summary>
    String,
    /// <summary>Decimal parsed using the request's explicit culture.</summary>
    Decimal,
    /// <summary>Signed 64-bit integer.</summary>
    Integer,
    /// <summary>True or false, case-insensitive.</summary>
    Boolean,
    /// <summary>Date parsed using an explicit exact format.</summary>
    Date
}

/// <summary>Named extraction requirement. Date fields require an explicit exact source format.</summary>
public sealed record OfficeAiFieldDefinition(string Name, OfficeAiFieldType Type = OfficeAiFieldType.String, string? DateFormat = null);

/// <summary>Document operation parameters. Collections are snapshotted before asynchronous work starts.</summary>
public sealed record OfficeAiRequest {
    /// <summary>Requested read-only operation.</summary>
    public OfficeAiOperation Operation { get; init; } = OfficeAiOperation.Ask;
    /// <summary>Question or extraction instruction, distinct from untrusted document evidence.</summary>
    public string Instruction { get; init; } = string.Empty;
    /// <summary>Explicit one-based pages, or an empty list for the whole captured source.</summary>
    public IReadOnlyList<int> Pages { get; init; } = Array.Empty<int>();
    /// <summary>Explicit evidence identifiers, or an empty list for all evidence in the page scope.</summary>
    public IReadOnlyList<string> EvidenceIds { get; init; } = Array.Empty<string>();
    /// <summary>Field requirements for ExtractFields.</summary>
    public IReadOnlyList<OfficeAiFieldDefinition> Fields { get; init; } = Array.Empty<OfficeAiFieldDefinition>();
    /// <summary>Culture for deterministic scalar normalization. Invariant culture is the default.</summary>
    public string Culture { get; init; } = string.Empty;
    /// <summary>Allows selected source images to leave the source boundary through the configured executor.</summary>
    public bool IncludeImages { get; init; }
    /// <summary>Explicit authorization to send selected evidence to a non-local execution profile.</summary>
    public bool AllowRemoteProcessing { get; init; }
    /// <summary>Whole-operation resource bounds.</summary>
    public OfficeAiLimits Limits { get; init; } = new();
}
