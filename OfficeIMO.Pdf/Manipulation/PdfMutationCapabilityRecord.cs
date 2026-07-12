namespace OfficeIMO.Pdf;

/// <summary>Input-specific support state for one shared PDF mutation capability family.</summary>
public sealed class PdfMutationCapabilityRecord {
    internal PdfMutationCapabilityRecord(
        PdfMutationCapabilityKind kind,
        IReadOnlyList<PdfMutationStructure> affectedStructures,
        bool fullRewriteImplemented,
        bool appendOnlyImplemented,
        bool fullRewriteAllowed,
        bool appendOnlyAllowed,
        IReadOnlyList<PdfMutationPermissionCheck> permissionChecks,
        IReadOnlyList<PdfMutationProof> requiredProofs,
        IReadOnlyList<string> blockerCodes) {
        Kind = kind;
        AffectedStructures = affectedStructures;
        FullRewriteImplemented = fullRewriteImplemented;
        AppendOnlyImplemented = appendOnlyImplemented;
        FullRewriteAllowed = fullRewriteAllowed;
        AppendOnlyAllowed = appendOnlyAllowed;
        PermissionChecks = permissionChecks;
        RequiredProofs = requiredProofs;
        BlockerCodes = blockerCodes;
    }

    /// <summary>Shared mutation capability family.</summary>
    public PdfMutationCapabilityKind Kind { get; }

    /// <summary>Structures in this capability family affected by the requested operation.</summary>
    public IReadOnlyList<PdfMutationStructure> AffectedStructures { get; }

    /// <summary>True when the engine implements a full-rewrite path for the requested operation.</summary>
    public bool FullRewriteImplemented { get; }

    /// <summary>True when the engine implements an append-only path for the requested operation.</summary>
    public bool AppendOnlyImplemented { get; }

    /// <summary>True when full rewrite is both implemented and permitted for this input.</summary>
    public bool FullRewriteAllowed { get; }

    /// <summary>True when append-only mutation is both implemented and permitted for this input.</summary>
    public bool AppendOnlyAllowed { get; }

    /// <summary>Permission and authorization checks relevant to the requested operation.</summary>
    public IReadOnlyList<PdfMutationPermissionCheck> PermissionChecks { get; }

    /// <summary>Evidence required when this operation is executed through its selected path.</summary>
    public IReadOnlyList<PdfMutationProof> RequiredProofs { get; }

    /// <summary>Stable blocker codes when the requested operation has no proven path.</summary>
    public IReadOnlyList<string> BlockerCodes { get; }
}
