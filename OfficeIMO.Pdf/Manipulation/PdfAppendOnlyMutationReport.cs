namespace OfficeIMO.Pdf;

/// <summary>
/// Describes append-only mutation support and blockers for an existing PDF.
/// </summary>
public sealed class PdfAppendOnlyMutationReport {
    internal PdfAppendOnlyMutationReport(
        PdfDocumentSecurityInfo security,
        IReadOnlyList<string> supportedActions,
        IReadOnlyList<string> blockedActions,
        IReadOnlyList<string> blockers,
        IReadOnlyList<string> warnings) {
        Security = security;
        SupportedActions = supportedActions;
        BlockedActions = blockedActions;
        Blockers = blockers;
        Warnings = warnings;
    }

    /// <summary>Security, signature, and revision markers used to decide append-only safety.</summary>
    public PdfDocumentSecurityInfo Security { get; }

    /// <summary>True when OfficeIMO.Pdf can append a metadata-only incremental revision to this input.</summary>
    public bool CanAppendMetadata => SupportedActions.Contains("Metadata", StringComparer.Ordinal);

    /// <summary>True when OfficeIMO.Pdf can append simple AcroForm field-value updates to this input.</summary>
    public bool CanAppendFormFields => SupportedActions.Contains("FormFill", StringComparer.Ordinal);

    /// <summary>True when OfficeIMO.Pdf can append an external signature placeholder revision to this input.</summary>
    public bool CanPrepareExternalSignature => SupportedActions.Contains("SignaturePrepare", StringComparer.Ordinal);

    /// <summary>True when the document structure permits append-only DSS/VRI enrichment.</summary>
    public bool CanAppendLongTermValidation => SupportedActions.Contains("LongTermValidation", StringComparer.Ordinal);

    /// <summary>True when the document structure and DocMDP policy permit append-only annotation changes.</summary>
    public bool CanAppendAnnotations => SupportedActions.Contains("Annotations", StringComparer.Ordinal);

    /// <summary>True when any append-only action can currently be applied by OfficeIMO.Pdf.</summary>
    public bool CanAppendAny => SupportedActions.Count > 0;

    /// <summary>True when the file markers indicate append-only mutation is required or preferred.</summary>
    public bool RequiresAppendOnlyMutation => Security.RequiresAppendOnlyMutation;

    /// <summary>True when OfficeIMO.Pdf must avoid all append-only mutation for this input.</summary>
    public bool BlocksAllAppendOnlyMutation => SupportedActions.Count == 0 && Blockers.Count > 0;

    /// <summary>Append-only actions currently supported for this input, for example Metadata.</summary>
    public IReadOnlyList<string> SupportedActions { get; }

    /// <summary>Append-only actions known to OfficeIMO.Pdf but blocked for this input or not implemented yet.</summary>
    public IReadOnlyList<string> BlockedActions { get; }

    /// <summary>Stable blocker codes explaining why append-only mutation is unavailable or limited.</summary>
    public IReadOnlyList<string> Blockers { get; }

    /// <summary>Non-blocking caution codes for automation workflows.</summary>
    public IReadOnlyList<string> Warnings { get; }

    /// <summary>Human-readable summary suitable for command-line surfaces.</summary>
    public string Summary {
        get {
            if (CanAppendAny) {
                return RequiresAppendOnlyMutation
                    ? "Incremental updates are supported for: " + string.Join(", ", SupportedActions) + "; other changes must remain append-only or be avoided."
                    : "Incremental updates are supported for: " + string.Join(", ", SupportedActions) + "; full rewrites may also be possible depending on preflight.";
            }

            return Blockers.Count == 0
                ? "No append-only actions are currently available for this input."
                : "Append-only mutation is blocked for this input: " + string.Join(", ", Blockers);
        }
    }
}
