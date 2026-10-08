using OfficeIMO.Email;

namespace OfficeIMO.Tool.Agent;

/// <summary>Compact mail-data metadata and optional HTML evidence. Cryptographic verification is never performed.</summary>
public sealed class AgentEmailInspectResult {
    /// <summary>Source-derived strings remain untrusted even when no finding is reported.</summary>
    public string ContentTrust => "untrusted";
    /// <summary>Selected-body evidence retained when detailed inspection output is trimmed.</summary>
    public AgentContentSafetySummary ContentSafety { get; set; } = new();
    public string SourceId { get; set; } = string.Empty;
    public string? Path { get; set; }
    public string Kind { get; set; } = string.Empty;
    public string Format { get; set; } = string.Empty;
    public string? ProtectionKind { get; set; }
    public string? SignatureStatus { get; set; }
    public int? BodyCount { get; set; }
    public int? AttachmentCount { get; set; }
    public int? ContainerCount { get; set; }
    public long? DeclaredItemCount { get; set; }
    public int? ContentLineRootCount { get; set; }
    public int DiagnosticCount { get; set; }
    public string HtmlInspectionStatus { get; set; } = string.Empty;
    public int? FindingSampleCount { get; set; }
    public bool? FindingLimitMayHaveBeenReached { get; set; }
    public int? BlockedElementCount { get; set; }
    public int? EventHandlerAttributeCount { get; set; }
    public EmailHtmlDataInspectionReport? Details { get; set; }
    public bool Truncated { get; set; }
}
