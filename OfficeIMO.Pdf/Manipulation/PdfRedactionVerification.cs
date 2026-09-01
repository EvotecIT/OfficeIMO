namespace OfficeIMO.Pdf;

/// <summary>
/// Provides reusable post-redaction checks for removed and retained PDF text markers.
/// </summary>
internal static partial class PdfRedactionVerification {
    /// <summary>
    /// Verifies a redacted PDF using the supplied redaction verification profile.
    /// </summary>
    public static PdfRedactionVerificationReport Verify(
        byte[] redactedPdf,
        PdfRedactionVerificationOptions options,
        PdfLoadOptions? readOptions = null) {
        Guard.NotNull(redactedPdf, nameof(redactedPdf));
        Guard.NotNull(options, nameof(options));

        PdfLoadOptions effectiveReadOptions = PdfLoadOptions.Resolve(readOptions);
        string extractedText = PdfReadDocument.Open(redactedPdf, effectiveReadOptions).ExtractText();
        string rawPdf = options.CheckRawPdfBytes ? PdfEncoding.Latin1GetString(redactedPdf) : string.Empty;
        var issues = new List<PdfRedactionVerificationIssue>();
        var externalResults = new List<PdfRedactionExternalValidationResult>();
        bool decodedPdfStreamsChecked = options.CheckDecodedPdfStreams || options.RequireCompleteStreamInspection;
        bool failOnUndecodablePdfStreams = options.FailOnUndecodablePdfStreams || options.RequireCompleteStreamInspection;

        if (decodedPdfStreamsChecked &&
            failOnUndecodablePdfStreams &&
            (options.RemovedTextMarkers.Count > 0 || options.RequireCompleteStreamInspection)) {
            issues.AddRange(FindUndecodableStreamIssues(redactedPdf, effectiveReadOptions));
        }

        for (int i = 0; i < options.RemovedTextMarkers.Count; i++) {
            string marker = options.RemovedTextMarkers[i];
            if (ContainsMarker(extractedText, marker, options.MatchCase)) {
                issues.Add(new PdfRedactionVerificationIssue(
                    "RemovedTextMarker",
                    marker,
                    "Removed text marker remains extractable after redaction: " + marker));
            }

            if (options.CheckRawPdfBytes && ContainsMarker(rawPdf, marker, options.MatchCase)) {
                issues.Add(new PdfRedactionVerificationIssue(
                    "RemovedRawMarker",
                    marker,
                    "Removed text marker remains in raw rewritten PDF bytes: " + marker));
            }

            if (options.CheckEncodedPdfStrings && ContainsEncodedPdfMarker(redactedPdf, marker, options.MatchCase)) {
                issues.Add(new PdfRedactionVerificationIssue(
                    "RemovedEncodedMarker",
                    marker,
                    "Removed text marker remains in encoded rewritten PDF string bytes: " + marker));
            }

            if (decodedPdfStreamsChecked && ContainsDecodedStreamMarker(redactedPdf, marker, options.MatchCase, effectiveReadOptions)) {
                issues.Add(new PdfRedactionVerificationIssue(
                    "RemovedDecodedStreamMarker",
                    marker,
                    "Removed text marker remains in a decoded rewritten PDF stream: " + marker));
            }
        }

        for (int i = 0; i < options.RetainedTextMarkers.Count; i++) {
            string marker = options.RetainedTextMarkers[i];
            if (!ContainsMarker(extractedText, marker, options.MatchCase)) {
                issues.Add(new PdfRedactionVerificationIssue(
                    "RetainedTextMarker",
                    marker,
                    "Expected retained text marker is not extractable after redaction: " + marker));
            }
        }

        if (options.CheckManagedRendering) {
            IReadOnlyList<PdfPageRenderResult> renders = PdfPageImageRenderer.RenderPages(redactedPdf, options: new PdfPageRenderOptions { Format = PdfPageRenderFormat.Svg, ContinueOnError = true }, readOptions: effectiveReadOptions);
            for (int i = 0; i < renders.Count; i++) if (!renders[i].Succeeded) issues.Add(new PdfRedactionVerificationIssue("ManagedRendering", renders[i].PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture), "Managed rendering failed for redacted page " + renders[i].PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + ": " + string.Join("; ", renders[i].Diagnostics)));
        }

        for (int i = 0; i < options.ExternalValidators.Count; i++) {
            PdfRedactionExternalValidationResult result = options.ExternalValidators[i].Validate((byte[])redactedPdf.Clone()); externalResults.Add(result);
            if (!result.IsValid) issues.Add(new PdfRedactionVerificationIssue("ExternalValidation", result.ValidatorName, "External redaction validation failed for " + result.ValidatorName + (string.IsNullOrWhiteSpace(result.Diagnostic) ? "." : ": " + result.Diagnostic)));
        }

        return new PdfRedactionVerificationReport(extractedText, options.CheckRawPdfBytes, options.CheckEncodedPdfStrings, decodedPdfStreamsChecked, options.CheckManagedRendering, externalResults.AsReadOnly(), issues.AsReadOnly());
    }

    /// <summary>
    /// Verifies configured markers and proves that the reviewed plan no longer finds text, image,
    /// or annotation intersections in the rewritten document.
    /// </summary>
    /// <remarks>
    /// The plan fingerprint binds review and application to the original source through
    /// <c>Apply(PdfRedactionPlan)</c>. This method verifies the supplied rewritten artifact's
    /// page shape and residual content; it does not independently prove rewrite lineage.
    /// </remarks>
    public static PdfRedactionVerificationReport VerifyAppliedPlan(
        byte[] redactedPdf,
        PdfRedactionPlan reviewedPlan,
        PdfRedactionVerificationOptions options,
        PdfLoadOptions? readOptions = null) {
        Guard.NotNull(reviewedPlan, nameof(reviewedPlan));
        PdfRedactionVerificationReport markerReport = Verify(redactedPdf, options, readOptions);
        PdfDocumentPreflight rewrittenPreflight = PdfInspector.Preflight(redactedPdf, readOptions);
        var issues = new List<PdfRedactionVerificationIssue>(markerReport.Issues);
        PdfDiagnosticFinding[] reviewedBlockingFindings = reviewedPlan.Findings
            .Where(static finding => finding.Severity == PdfDiagnosticSeverity.Error)
            .ToArray();
        if (!reviewedPlan.Preflight.CanReadLogicalObjects || reviewedBlockingFindings.Length > 0) {
            string detail = reviewedBlockingFindings.Length == 0
                ? string.Join(" ", reviewedPlan.Preflight.GetCapabilityDiagnostics(PdfPreflightCapability.ReadLogicalObjects))
                : string.Join(" ", reviewedBlockingFindings.Select(static finding => finding.Message));
            issues.Add(new PdfRedactionVerificationIssue(
                "ReviewedRedactionPlanBlocked",
                "ReviewedSource",
                "The original redaction plan was blocked and cannot provide redaction proof." +
                (string.IsNullOrWhiteSpace(detail) ? string.Empty : " " + detail)));
        }
        int? reviewedPageCount = reviewedPlan.Preflight.UncheckedDocumentInfo?.PageCount;
        int? rewrittenPageCount = rewrittenPreflight.UncheckedDocumentInfo?.PageCount;
        bool pageIdentityMatches = true;
        if (reviewedPageCount.HasValue &&
            rewrittenPageCount.HasValue &&
            reviewedPageCount.Value != rewrittenPageCount.Value) {
            pageIdentityMatches = false;
            issues.Add(new PdfRedactionVerificationIssue(
                "RedactionPlanPageCountChanged",
                "ReviewedPages",
                $"The reviewed PDF had {reviewedPageCount.Value} page(s), but the rewritten PDF has {rewrittenPageCount.Value}. Redaction verification requires the reviewed page set to be preserved."));
        }

        if (pageIdentityMatches &&
            reviewedPlan.PageIdentities.Count > 0 &&
            rewrittenPreflight.CanReadLogicalObjects) {
            IReadOnlyList<string> rewrittenPageIdentities = PdfRedactionPlan.CapturePageIdentities(PdfReadDocument.Open(redactedPdf, readOptions));
            if (!reviewedPlan.PageIdentities.SequenceEqual(rewrittenPageIdentities, StringComparer.Ordinal)) {
                pageIdentityMatches = false;
                issues.Add(new PdfRedactionVerificationIssue(
                    "RedactionPlanPageIdentityChanged",
                    "ReviewedPages",
                    "The rewritten PDF changed the reviewed page order, rotation, MediaBox, CropBox, or UserUnit. Redaction verification will not reuse reviewed rectangles in a different page coordinate system."));
            }
        }

        if (rewrittenPageCount.HasValue) {
            foreach (int missingPageNumber in reviewedPlan.Areas
                .Select(static area => area.PageNumber)
                .Where(pageNumber => pageNumber < 1 || pageNumber > rewrittenPageCount.Value)
                .Distinct()
                .OrderBy(static pageNumber => pageNumber)) {
                issues.Add(new PdfRedactionVerificationIssue(
                    "RedactionPlanPageMissing",
                    "Page:" + missingPageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    $"Reviewed redaction page {missingPageNumber} does not exist in the rewritten PDF."));
            }
        }

        PdfRedactionPlan? residualPlan = reviewedPlan.Areas.Count == 0 || !pageIdentityMatches
            ? null
            : PdfRedactionPlanner.Plan(redactedPdf, reviewedPlan.Areas, options: readOptions);

        PdfDiagnosticFinding[] blockingFindings = (residualPlan?.Findings ?? Array.Empty<PdfDiagnosticFinding>())
            .Where(static finding => finding.Severity == PdfDiagnosticSeverity.Error)
            .ToArray();
        if (!rewrittenPreflight.CanReadLogicalObjects || blockingFindings.Length > 0) {
            string detail = blockingFindings.Length == 0
                ? string.Join(" ", rewrittenPreflight.GetCapabilityDiagnostics(PdfPreflightCapability.ReadLogicalObjects))
                : string.Join(" ", blockingFindings.Select(static finding => finding.Message));
            issues.Add(new PdfRedactionVerificationIssue(
                "RedactionPlanInspectionBlocked",
                "ReviewedAreas",
                "The rewritten PDF could not be inspected for residual content inside the reviewed redaction areas." +
                (string.IsNullOrWhiteSpace(detail) ? string.Empty : " " + detail)));
        }

        foreach (IGrouping<(PdfRedactionMatchKind Kind, int PageNumber), PdfRedactionMatch> group in (residualPlan?.Matches ?? Array.Empty<PdfRedactionMatch>())
            .GroupBy(static match => (match.Kind, match.PageNumber))) {
            string marker = group.Key.Kind + "@page:" + group.Key.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture);
            issues.Add(new PdfRedactionVerificationIssue(
                "RedactionPlanResidual",
                marker,
                $"The rewritten PDF still contains {group.Count()} {group.Key.Kind} match(es) inside a reviewed redaction area on page {group.Key.PageNumber}."));
        }

        if (issues.Count == markerReport.Issues.Count) {
            return markerReport;
        }

        return new PdfRedactionVerificationReport(
            markerReport.ExtractedText,
            markerReport.RawPdfBytesChecked,
            markerReport.EncodedPdfStringsChecked,
            markerReport.DecodedPdfStreamsChecked,
            markerReport.ManagedRenderingChecked,
            markerReport.ExternalValidationResults,
            issues.AsReadOnly());
    }

    /// <summary>
    /// Verifies a redacted PDF and throws when removed text remains or retained text disappears.
    /// </summary>
    public static PdfRedactionVerificationReport AssertVerified(
        byte[] redactedPdf,
        PdfRedactionVerificationOptions options,
        PdfLoadOptions? readOptions = null) {
        PdfRedactionVerificationReport report = Verify(redactedPdf, options, readOptions);
        report.ThrowIfFailed();
        return report;
    }

    /// <summary>Verifies a reviewed redaction plan and throws when any planned content class remains in its areas.</summary>
    public static PdfRedactionVerificationReport AssertAppliedPlan(
        byte[] redactedPdf,
        PdfRedactionPlan reviewedPlan,
        PdfRedactionVerificationOptions options,
        PdfLoadOptions? readOptions = null) {
        PdfRedactionVerificationReport report = VerifyAppliedPlan(redactedPdf, reviewedPlan, options, readOptions);
        report.ThrowIfFailed();
        return report;
    }

    private static bool ContainsMarker(string text, string marker, bool matchCase) {
        if (string.IsNullOrEmpty(marker)) {
            return false;
        }

#if NETSTANDARD2_0 || NETFRAMEWORK
        return text.IndexOf(
            marker,
            matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase) >= 0;
#else
        return text.Contains(
            marker,
            matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase);
#endif
    }
}
