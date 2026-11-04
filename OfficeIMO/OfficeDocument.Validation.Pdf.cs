using OfficeIMO.Pdf;

namespace OfficeIMO;

public static partial class OfficeDocument {
    private static PdfLoadOptions ValidationPdfRead(OfficeDocumentValidationOptions options, PdfDocument? document = null) {
        PdfLoadOptions? modelPolicy = document?.ReadOptions;
        PdfLoadOptions policy = options.PdfRead ?? options.Pdf?.Rewrite?.OriginalReadOptions ?? modelPolicy ?? new PdfLoadOptions();
        return policy.Clone(null, options.Password ?? policy.Password ?? modelPolicy?.Password);
    }

    private static void ValidatePdf(PdfDocument document, OfficeDocumentCheck check, OfficeDocumentValidationOptions options,
        PdfLoadOptions originalRead, List<OfficeDiagnostic> diagnostics, List<object> evidence, CancellationToken cancellationToken, string? sourcePath) {
        var settings = options.Pdf ?? new OfficePdfValidationOptions();
        switch (check) {
            case OfficeDocumentCheck.Structure:
            case OfficeDocumentCheck.Security:
                var preflight = document.Preflight(originalRead, cancellationToken);
                evidence.Add(preflight);
                foreach (var blocker in preflight.ReadBlockers) diagnostics.Add(new OfficeDiagnostic("Pdf.Read." + blocker.Kind,
                    OfficeDiagnosticSeverity.Error, blocker.Message, format: "PDF"));
                foreach (var blocker in preflight.RewriteBlockers) diagnostics.Add(new OfficeDiagnostic("Pdf.Rewrite." + blocker.Kind,
                    OfficeDiagnosticSeverity.Warning, blocker.Message, format: "PDF"));
                foreach (string message in preflight.Diagnostics) diagnostics.Add(new OfficeDiagnostic("Pdf.Preflight",
                    OfficeDiagnosticSeverity.Information, message, format: "PDF"));
                if (!preflight.CanRead && preflight.ReadBlockers.Count == 0)
                    diagnostics.Add(new OfficeDiagnostic("Pdf.Unreadable", OfficeDiagnosticSeverity.Error,
                        "The native PDF preflight could not read this document.", format: "PDF"));
                if (check == OfficeDocumentCheck.Security) {
                    ValidationNotApplicable(diagnostics, check, "PDF",
                        "The Office package-security policy applies to ZIP and compound containers. Native PDF security and permission evidence is retained in NativeReports.");
                    break;
                }
                foreach (var capability in settings.Capability ?? Array.Empty<PdfPreflightCapability>()) {
                    if (!preflight.Can(capability)) diagnostics.Add(new OfficeDiagnostic("Pdf.Capability." + capability,
                        OfficeDiagnosticSeverity.Error, string.Join("; ", preflight.GetCapabilityDiagnostics(capability)), format: "PDF"));
                }
                if (settings.DifferenceDocument != null || settings.DifferencePath != null) {
                    cancellationToken.ThrowIfCancellationRequested();
                    var preservation = ValidationPdfPreservation(settings.Rewrite, originalRead);
                    PdfRewritePreservationReport rewrite;
                    if (settings.DifferenceDocument != null) rewrite = document.Proof.AssessRewritePreservation(settings.DifferenceDocument, preservation);
                    else {
                        var difference = PdfDocument.Load(settings.DifferencePath!, preservation.RewrittenReadOptions, cancellationToken);
                        try { rewrite = document.Proof.AssessRewritePreservation(difference, preservation); }
                        finally { (difference as IDisposable)?.Dispose(); }
                    }
                    evidence.Add(rewrite);
                    foreach (var issue in rewrite.Issues) diagnostics.Add(new OfficeDiagnostic("Pdf.Preservation." + issue.Feature,
                        OfficeDiagnosticSeverity.Error, issue.Message, sourcePath, "PDF"));
                }
                break;
            case OfficeDocumentCheck.Compliance:
            case OfficeDocumentCheck.Accessibility:
                PdfComplianceProfile profile = check == OfficeDocumentCheck.Accessibility ? PdfComplianceProfile.PdfUa1
                    : settings.Profile ?? document.AssessCompliance().Profile;
                if (profile == PdfComplianceProfile.None) {
                    ValidationNotApplicable(diagnostics, check, "PDF", "Select a formal Profile in PDF validation options.");
                    break;
                }
                if (settings.Proof || settings.ExternalValidation != null) {
                    var proof = document.CreateComplianceArtifact(profile).AssessProof(settings.ExternalValidation);
                    evidence.Add(proof);
                    AddPdfValidationRequirements(proof.EffectiveRequirements, diagnostics);
                    foreach (var validator in proof.ExternalValidatorProofs) {
                        if (!validator.IsSatisfied) diagnostics.Add(new OfficeDiagnostic("Pdf.ExternalValidation",
                            OfficeDiagnosticSeverity.Error, validator.Diagnostic, format: "PDF"));
                    }
                } else {
                    var readiness = document.AssessCompliance(profile, cancellationToken);
                    evidence.Add(readiness);
                    AddPdfValidationRequirements(readiness.Requirements, diagnostics);
                }
                break;
            default:
                ValidationNotApplicable(diagnostics, check, "PDF");
                break;
        }
        cancellationToken.ThrowIfCancellationRequested();
    }

    // Preserve the native signals and exclusions while replacing only the original side with its captured policy.
    // Native preservation options have mutable collections and no clone API; neither they nor their read policies are borrowed for mutation.
    private static PdfRewritePreservationOptions ValidationPdfPreservation(PdfRewritePreservationOptions? source, PdfLoadOptions originalRead) {
        source ??= new PdfRewritePreservationOptions();
        var result = new PdfRewritePreservationOptions {
            OriginalReadOptions = originalRead,
            RewrittenReadOptions = source.RewrittenReadOptions?.Clone(),
            PreservePageCount = source.PreservePageCount,
            PreservePageGeometry = source.PreservePageGeometry,
            PreserveMetadata = source.PreserveMetadata,
            PreserveOutlines = source.PreserveOutlines,
            PreserveNamedDestinations = source.PreserveNamedDestinations,
            PreservePageLabels = source.PreservePageLabels,
            PreserveLinkAnnotations = source.PreserveLinkAnnotations,
            PreserveAnnotations = source.PreserveAnnotations,
            PreserveForms = source.PreserveForms,
            PreserveEmbeddedFiles = source.PreserveEmbeddedFiles,
            PreserveXmpMetadata = source.PreserveXmpMetadata,
            PreserveOutputIntents = source.PreserveOutputIntents,
            PreserveOptionalContent = source.PreserveOptionalContent,
            PreserveTaggedContent = source.PreserveTaggedContent,
            PreserveCatalogViewSettings = source.PreserveCatalogViewSettings,
            PreserveOpenAction = source.PreserveOpenAction,
            PreserveCatalogActions = source.PreserveCatalogActions,
            PreservePageActions = source.PreservePageActions,
            PreserveFormWidgetActions = source.PreserveFormWidgetActions,
            FilterActionsByPreservedTypes = source.FilterActionsByPreservedTypes,
            PreserveViewerPreferences = source.PreserveViewerPreferences,
            PreserveDocumentVersionState = source.PreserveDocumentVersionState,
            PreserveRevisionStructure = source.PreserveRevisionStructure,
            PreserveSecurityState = source.PreserveSecurityState
        };
        foreach (string marker in source.RequiredTextMarkers) result.RequiredTextMarkers.Add(marker);
        result.AllowedMetadataChanges.UnionWith(source.AllowedMetadataChanges);
        result.PreservedActionTypes.UnionWith(source.PreservedActionTypes);
        result.ExcludedAnnotationSubtypes.UnionWith(source.ExcludedAnnotationSubtypes);
        result.ExcludedLinkAnnotationUris.UnionWith(source.ExcludedLinkAnnotationUris);
        result.ExcludedActionUris.UnionWith(source.ExcludedActionUris);
        return result;
    }

    private static void AddPdfValidationRequirements(IReadOnlyList<PdfComplianceRequirement> requirements, List<OfficeDiagnostic> diagnostics) {
        foreach (var requirement in requirements) diagnostics.Add(new OfficeDiagnostic("Pdf.Compliance." + requirement.Id,
            requirement.Status == PdfComplianceRequirementStatus.Satisfied ? OfficeDiagnosticSeverity.Information : OfficeDiagnosticSeverity.Error,
            requirement.Diagnostic, format: "PDF"));
    }
}
