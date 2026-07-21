namespace OfficeIMO.Pdf;

/// <summary>Explains how OfficeIMO.Pdf should execute and prove an existing-document mutation.</summary>
public sealed class PdfMutationPlan {
    internal PdfMutationPlan(
        PdfMutationOperation operation,
        PdfMutationExecutionPreference executionPreference,
        PdfMutationExecutionMode executionMode,
        PdfDocumentPreflight preflight,
        PdfAppendOnlyMutationReport appendOnlyReport,
        bool fullRewriteAvailable,
        bool appendOnlyAvailable,
        IReadOnlyList<PdfMutationStructure> affectedStructures,
        IReadOnlyList<PdfMutationCapabilityRecord> capabilityRecords,
        IReadOnlyList<PdfMutationPermissionCheck> permissionChecks,
        IReadOnlyList<PdfMutationProof> requiredProofs,
        IReadOnlyList<string> blockerCodes,
        IReadOnlyList<string> warnings,
        IReadOnlyList<string> diagnostics) {
        Operation = operation;
        ExecutionPreference = executionPreference;
        ExecutionMode = executionMode;
        Preflight = preflight;
        AppendOnlyReport = appendOnlyReport;
        FullRewriteAvailable = fullRewriteAvailable;
        AppendOnlyAvailable = appendOnlyAvailable;
        AffectedStructures = affectedStructures;
        CapabilityRecords = capabilityRecords;
        PermissionChecks = permissionChecks;
        RequiredProofs = requiredProofs;
        BlockerCodes = blockerCodes;
        Warnings = warnings;
        Diagnostics = diagnostics;
    }

    /// <summary>Requested mutation family.</summary>
    public PdfMutationOperation Operation { get; }

    /// <summary>Caller constraint used when selecting the execution mode.</summary>
    public PdfMutationExecutionPreference ExecutionPreference { get; }

    /// <summary>Selected execution mode.</summary>
    public PdfMutationExecutionMode ExecutionMode { get; }

    /// <summary>True when the current engine has a proven execution path for this request.</summary>
    public bool CanExecute => ExecutionMode != PdfMutationExecutionMode.Blocked;

    /// <summary>General read and rewrite preflight used to build this plan.</summary>
    public PdfDocumentPreflight Preflight { get; }

    /// <summary>Append-only policy used to build this plan.</summary>
    public PdfAppendOnlyMutationReport AppendOnlyReport { get; }

    /// <summary>True when a full rewrite is currently permitted for this operation and input.</summary>
    public bool FullRewriteAvailable { get; }

    /// <summary>True when an append-only implementation is currently available for this operation and input.</summary>
    public bool AppendOnlyAvailable { get; }

    /// <summary>PDF structures the requested operation can affect.</summary>
    public IReadOnlyList<PdfMutationStructure> AffectedStructures { get; }

    /// <summary>Input-specific shared capability records that contributed to this decision.</summary>
    public IReadOnlyList<PdfMutationCapabilityRecord> CapabilityRecords { get; }

    /// <summary>Permission and authorization rules that must be evaluated for this operation.</summary>
    public IReadOnlyList<PdfMutationPermissionCheck> PermissionChecks { get; }

    /// <summary>Evidence required after the mutation is applied.</summary>
    public IReadOnlyList<PdfMutationProof> RequiredProofs { get; }

    /// <summary>Stable machine-readable blocker codes when no execution path is available.</summary>
    public IReadOnlyList<string> BlockerCodes { get; }

    /// <summary>Stable caution codes that do not block the selected execution path.</summary>
    public IReadOnlyList<string> Warnings { get; }

    /// <summary>Human-readable explanation of the selected or blocked plan.</summary>
    public IReadOnlyList<string> Diagnostics { get; }

    /// <summary>True when this executable plan explicitly ignores authenticated user-password permission restrictions.</summary>
    public bool PermissionRestrictionsIgnored => CanExecute && Preflight.PermissionRestrictionsIgnored;

    /// <summary>Short plan summary suitable for logs and command surfaces.</summary>
    public string Summary => CanExecute
        ? Operation + " will use " + ExecutionMode + " and requires " + RequiredProofs.Count + " proof check(s)."
        : Operation + " is blocked: " + (BlockerCodes.Count == 0 ? "no proven execution path." : string.Join(", ", BlockerCodes) + ".");
}
