namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentRedactions {
    /// <summary>
    /// Applies a reviewed redaction, sanitizes the rewritten artifact using an explicit policy,
    /// and rechecks the final artifact before returning it. No source file is modified.
    /// </summary>
    /// <remarks>The final verification checks the original reviewed areas and configured markers again,
    /// because sanitization is a second rewrite. Report summaries omit document text and search criteria.</remarks>
    public PdfRedactionSharingResult ApplyForSharing(
        PdfRedactionPlan plan,
        PdfSanitizationOptions sanitizationOptions,
        PdfRedactionApplyOptions? applyOptions = null,
        PdfRedactionVerificationOptions? verificationOptions = null) {
        Guard.NotNull(sanitizationOptions, nameof(sanitizationOptions));
        sanitizationOptions.CancellationToken.ThrowIfCancellationRequested();
        var verification = verificationOptions ?? new PdfRedactionVerificationOptions {
            CheckManagedRendering = true, RequireCompleteStreamInspection = true, FailOnUndecodablePdfStreams = true,
            CancellationToken = sanitizationOptions.CancellationToken
        };
        applyOptions ??= new PdfRedactionApplyOptions { CancellationToken = sanitizationOptions.CancellationToken };
        PdfRedactionApplyResult redaction = ApplyWithEvidence(plan, applyOptions, verification).ThrowIfUnverified();
        sanitizationOptions.CancellationToken.ThrowIfCancellationRequested();
        PdfSanitizationResult sanitization = redaction.ToDocument().Sanitize(sanitizationOptions);
        if (!sanitization.IsSanitized || !sanitization.PreservationReport.IsPreserved) {
            throw new InvalidOperationException("The redacted artifact did not pass sanitization and preservation checks.");
        }
        sanitizationOptions.CancellationToken.ThrowIfCancellationRequested();
        PdfDocument finalDocument = sanitization.ToDocument();
        byte[] finalBytes = sanitization.ToBytes();
        IReadOnlyList<string> beforeContent = PdfRedactionPlan.CapturePageContentIdentities(
            PdfReadDocument.Open(redaction.Pdf, _document.ReadOptions, sanitizationOptions.CancellationToken));
        IReadOnlyList<string> afterContent = PdfRedactionPlan.CapturePageContentIdentities(
            PdfReadDocument.Open(finalBytes, finalDocument.ReadOptions, sanitizationOptions.CancellationToken));
        if (!beforeContent.SequenceEqual(afterContent, StringComparer.Ordinal)) {
            throw new InvalidOperationException("Sanitization changed page content or geometry after redaction; the combined artifact cannot be verified.");
        }
        // The ordinary plan identity also includes annotations and links that sanitization may
        // intentionally remove. Its replacement is the full page-content comparison above plus
        // the sanitizer's policy-specific preservation report, not a blanket identity bypass.
        var contentPlan = new PdfRedactionPlan(plan.Preflight, plan.Areas, plan.Matches, plan.Findings,
            plan.SearchCriteria, plan.SourceSha256, reviewedTextObjectScopes: plan.ReviewedTextObjectScopes);
        PdfRedactionMatch[] provenImageRewrites = redaction.Evidence.Items
            .Where(item => item.Status == PdfRedactionEvidenceStatus.VerifiedAbsent && item.ReviewedMatch.Kind == PdfRedactionMatchKind.ImagePlacement)
            .Select(item => item.ReviewedMatch).ToArray();
        PdfRedactionVerificationReport finalVerification = PdfRedactionVerification.VerifyAppliedPlan(
            finalBytes, contentPlan, verification, finalDocument.ReadOptions, provenImageRewrites);
        finalVerification.ThrowIfFailed();
        sanitizationOptions.CancellationToken.ThrowIfCancellationRequested();
        return new PdfRedactionSharingResult(redaction.Evidence, sanitization, finalVerification);
    }
}
