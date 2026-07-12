namespace OfficeIMO.Pdf;

/// <summary>Before/after signature and revision proof for a requested PDF mutation.</summary>
public sealed class PdfSignatureMutationReport {
    internal PdfSignatureMutationReport(
        PdfMutationPlan mutationPlan,
        PdfSignatureValidationReport before,
        PdfSignatureValidationReport after,
        bool originalBytesArePrefix,
        bool revisionChainExtended,
        IReadOnlyList<PdfSignatureMutationResult> signatures,
        IReadOnlyList<string> diagnostics) {
        MutationPlan = mutationPlan;
        Before = before;
        After = after;
        OriginalBytesArePrefix = originalBytesArePrefix;
        RevisionChainExtended = revisionChainExtended;
        Signatures = signatures;
        Diagnostics = diagnostics;
    }

    /// <summary>Mutation decision and DocMDP/FieldMDP policy evaluated against the input.</summary>
    public PdfMutationPlan MutationPlan { get; }

    /// <summary>Signature structure before mutation.</summary>
    public PdfSignatureValidationReport Before { get; }

    /// <summary>Signature structure after mutation.</summary>
    public PdfSignatureValidationReport After { get; }

    /// <summary>True when every input byte remains an exact prefix of the output.</summary>
    public bool OriginalBytesArePrefix { get; }

    /// <summary>True when the output adds a revision whose /Prev points to the input's final xref revision.</summary>
    public bool RevisionChainExtended { get; }

    /// <summary>Per-signature coverage and preservation results.</summary>
    public IReadOnlyList<PdfSignatureMutationResult> Signatures { get; }

    /// <summary>Stable structural proof diagnostics.</summary>
    public IReadOnlyList<string> Diagnostics { get; }

    /// <summary>True when the mutation planner permits the requested change.</summary>
    public bool RequestedChangeIsPermitted => MutationPlan.CanExecute;

    /// <summary>True when all pre-existing signatures remain structurally preserved.</summary>
    public bool AllExistingSignaturesArePreserved => Signatures.All(static signature => signature.IsStructurallyPreserved);

    /// <summary>True when append-only and signature structural proof both pass.</summary>
    public bool IsPreservedAppendOnlyMutation =>
        MutationPlan.ExecutionMode == PdfMutationExecutionMode.AppendOnly &&
        OriginalBytesArePrefix &&
        RevisionChainExtended &&
        AllExistingSignaturesArePreserved;
}
