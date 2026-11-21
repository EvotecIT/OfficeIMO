namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentRedactions {
    /// <summary>
    /// Applies a reviewed redaction, removes the selected hidden data from the rewritten artifact,
    /// and rechecks the final artifact before returning it. No source file is modified.
    /// </summary>
    /// <remarks>The final verification checks the original reviewed areas and configured markers again,
    /// because hidden-data removal is a second rewrite. Configured text markers and external validators check
    /// the final artifact. Report summaries omit document text and search criteria.</remarks>
    public PdfRedactionSharingResult ApplyForSharing(
        PdfRedactionPlan plan,
        OfficeHiddenDataOptions hiddenDataOptions,
        PdfRedactionApplyOptions? applyOptions = null,
        PdfRedactionVerificationOptions? verificationOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        Guard.NotNull(hiddenDataOptions, nameof(hiddenDataOptions));
        applyOptions = PdfOperationCancellation.Bind(applyOptions, createWhenNull: true, cancellationToken);
        verificationOptions = PdfOperationCancellation.Bind(verificationOptions, createWhenNull: true, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        var verification = verificationOptions ?? new PdfRedactionVerificationOptions {
            CheckManagedRendering = true, RequireCompleteStreamInspection = true, FailOnUndecodablePdfStreams = true,
            CancellationToken = cancellationToken
        };
        applyOptions ??= new PdfRedactionApplyOptions { CancellationToken = cancellationToken };
        // Global markers can remain in metadata or attachments until the requested hidden-data removal runs.
        // Prove the reviewed page-content rewrite first; apply whole-artifact requirements only
        // to the final output so validators never receive the intermediate artifact.
        PdfRedactionApplyResult redaction = ApplyWithEvidence(plan, applyOptions,
            CreateIntermediateVerification(verification), cancellationToken: cancellationToken).ThrowIfUnverified();
        cancellationToken.ThrowIfCancellationRequested();
        PdfDocument redactedDocument = redaction.ToDocument();
        PdfHiddenDataRemoval removal = redactedDocument.RemoveHiddenDataCore(hiddenDataOptions, maximumOutputBytes: null, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        byte[] finalBytes = removal.Bytes;
        IReadOnlyList<string> beforeContent = PdfRedactionPlan.CapturePageContentIdentities(
            PdfReadDocument.Open(redaction.Pdf, redactedDocument.ReadOptions, cancellationToken));
        IReadOnlyList<string> afterContent = PdfRedactionPlan.CapturePageContentIdentities(
            PdfReadDocument.Open(finalBytes, removal.ReadOptions, cancellationToken));
        if (!beforeContent.SequenceEqual(afterContent, StringComparer.Ordinal)) {
            throw new InvalidOperationException("Hidden-data removal changed page content or geometry after redaction; the combined artifact cannot be verified.");
        }
        // The ordinary plan identity also includes annotations and links that hidden-data removal may
        // intentionally remove. Its replacement is the full page-content comparison above plus
        // the engine's selection-specific preservation report, not a blanket identity bypass.
        var contentPlan = new PdfRedactionPlan(plan.Preflight, plan.Areas, plan.Matches, plan.Findings,
            plan.RequestedSearchCriteria, plan.SourceSha256, reviewedTextObjectScopes: plan.ReviewedTextObjectScopes,
            searchMatchCase: plan.SearchMatchCase, searchRegexOptions: plan.SearchRegexOptions,
            searchRegexTimeout: plan.SearchRegexTimeout, searchPageNumbers: plan.SearchPageNumbers, alternateTextSections: plan.AlternateTextSections, alternateTextPlacements: plan.AlternateTextPlacements);
        PdfRedactionMatch[] provenImageRewrites = redaction.Evidence.Items
            .Where(item => item.Status == PdfRedactionEvidenceStatus.VerifiedAbsent && item.ReviewedMatch.Kind == PdfRedactionMatchKind.ImagePlacement)
            .Select(item => item.ReviewedMatch).ToArray();
        PdfRedactionVerificationReport finalVerification = PdfRedactionVerification.VerifyAppliedPlan(
            finalBytes, contentPlan, verification, removal.ReadOptions, provenImageRewrites);
        finalVerification.ThrowIfFailed();
        cancellationToken.ThrowIfCancellationRequested();
        return new PdfRedactionSharingResult(redaction.Evidence, removal, finalVerification);
    }

    private static PdfRedactionVerificationOptions CreateIntermediateVerification(PdfRedactionVerificationOptions final) => new() {
        CheckRawPdfBytes = final.CheckRawPdfBytes,
        CheckEncodedPdfStrings = final.CheckEncodedPdfStrings,
        CheckDecodedPdfStreams = final.CheckDecodedPdfStreams,
        FailOnUndecodablePdfStreams = final.FailOnUndecodablePdfStreams,
        RequireCompleteStreamInspection = final.RequireCompleteStreamInspection,
        MatchCase = final.MatchCase,
        CheckManagedRendering = final.CheckManagedRendering,
        CancellationToken = final.CancellationToken
    };
}
