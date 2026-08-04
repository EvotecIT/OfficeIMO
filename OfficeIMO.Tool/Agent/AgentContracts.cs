namespace OfficeIMO.Tool.Agent;

/// <summary>Compact description of one local OfficeIMO-readable artifact.</summary>
public sealed class AgentInspectResult {
    public string SourceId { get; set; } = string.Empty;
    public string Path { get; set; } = string.Empty;
    public string Kind { get; set; } = string.Empty;
    public string? Format { get; set; }
    public long? LengthBytes { get; set; }
    public string? Title { get; set; }
    public string? Author { get; set; }
    public string? Subject { get; set; }
    public string? Preview { get; set; }
    public int ChunkCount { get; set; }
    public int BlockCount { get; set; }
    public int PageCount { get; set; }
    public int TableCount { get; set; }
    public int AssetCount { get; set; }
    public int MetadataCount { get; set; }
    public int DiagnosticCount { get; set; }
    public int FolderCount { get; set; }
    public IReadOnlyList<AgentFolderSummary> Folders { get; set; } = Array.Empty<AgentFolderSummary>();
    public IReadOnlyList<AgentMetadataSummary> Metadata { get; set; } = Array.Empty<AgentMetadataSummary>();
    public IReadOnlyList<AgentDiagnosticSummary> Diagnostics { get; set; } = Array.Empty<AgentDiagnosticSummary>();
    public bool Truncated { get; set; }
}

/// <summary>Compact folder metadata for email-store inspection.</summary>
public sealed class AgentFolderSummary {
    public string Id { get; set; } = string.Empty;
    public string? ParentId { get; set; }
    public string Name { get; set; } = string.Empty;
    public int? ItemCount { get; set; }
    public int? AssociatedItemCount { get; set; }
    public string? SpecialKind { get; set; }
}

/// <summary>Compact scalar metadata projected from a document or email item.</summary>
public sealed class AgentMetadataSummary {
    public string Name { get; set; } = string.Empty;
    public string? Value { get; set; }
}

/// <summary>Compact diagnostic sample with stable severity and code.</summary>
public sealed class AgentDiagnosticSummary {
    public string Code { get; set; } = string.Empty;
    public string Severity { get; set; } = string.Empty;
    public string Message { get; set; } = string.Empty;
}

/// <summary>Bounded search response for a document or email store.</summary>
public sealed class AgentSearchResult {
    public string SourceId { get; set; } = string.Empty;
    public string? Query { get; set; }
    public int Returned { get; set; }
    public int? NextCursor { get; set; }
    public bool Truncated { get; set; }
    public IReadOnlyList<AgentSearchHit> Results { get; set; } = Array.Empty<AgentSearchHit>();
}

/// <summary>One compact result that can be retrieved with the fetch operation.</summary>
public sealed class AgentSearchHit {
    public string Id { get; set; } = string.Empty;
    public string? Title { get; set; }
    public string? Snippet { get; set; }
    public string? Sender { get; set; }
    public DateTimeOffset? Timestamp { get; set; }
    public string? FolderId { get; set; }
    public IReadOnlyList<int>? Pages { get; set; }
}

/// <summary>Bounded content retrieved for one opaque result identifier.</summary>
public sealed class AgentFetchResult {
    public string SourceId { get; set; } = string.Empty;
    public string Id { get; set; } = string.Empty;
    public string Kind { get; set; } = string.Empty;
    public string? Title { get; set; }
    public string Content { get; set; } = string.Empty;
    public int ContentLength { get; set; }
    public int? NextCursor { get; set; }
    public bool Truncated { get; set; }
    public IReadOnlyList<AgentMetadataSummary> Metadata { get; set; } = Array.Empty<AgentMetadataSummary>();
    public IReadOnlyList<AgentDiagnosticSummary> Diagnostics { get; set; } = Array.Empty<AgentDiagnosticSummary>();
}

/// <summary>Filtered, token-efficient OfficeIMO Reader capability response.</summary>
public sealed class AgentCapabilitiesResult {
    public string? Extension { get; set; }
    public string Operation { get; set; } = "read";
    public int Returned { get; set; }
    public bool Truncated { get; set; }
    public IReadOnlyList<AgentCapabilitySummary> Capabilities { get; set; } = Array.Empty<AgentCapabilitySummary>();
    public int ConversionReturned { get; set; }
    public IReadOnlyList<AgentConversionCapabilitySummary> Conversions { get; set; } = Array.Empty<AgentConversionCapabilitySummary>();
}

/// <summary>Small capability description intended for CLI and MCP discovery.</summary>
public sealed class AgentCapabilitySummary {
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string Kind { get; set; } = string.Empty;
    public IReadOnlyList<string> Extensions { get; set; } = Array.Empty<string>();
}

/// <summary>Small package-neutral conversion route intended for CLI and MCP discovery.</summary>
public sealed class AgentConversionCapabilitySummary {
    public string Id { get; set; } = string.Empty;
    public string Source { get; set; } = string.Empty;
    public string Target { get; set; } = string.Empty;
    public IReadOnlyList<string> SourceExtensions { get; set; } = Array.Empty<string>();
    public string TargetExtension { get; set; } = string.Empty;
    public string PackageId { get; set; } = string.Empty;
    public string Fidelity { get; set; } = string.Empty;
    public string ResultContract { get; set; } = string.Empty;
    public bool BrowserAvailable { get; set; }
}

/// <summary>Result of materializing one OfficeIMO Reader representation.</summary>
public sealed class AgentConvertResult {
    public string SourceId { get; set; } = string.Empty;
    public string SourcePath { get; set; } = string.Empty;
    public string OutputPath { get; set; } = string.Empty;
    public string Format { get; set; } = string.Empty;
    public long LengthBytes { get; set; }
    public string Sha256 { get; set; } = string.Empty;
    public int DiagnosticCount { get; set; }
}
