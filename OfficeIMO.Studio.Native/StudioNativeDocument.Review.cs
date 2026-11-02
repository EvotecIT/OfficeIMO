using System.Globalization;
using OfficeIMO.Pdf;
using OfficeIMO.Security;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Native;

/// <summary>
/// Reviewing a document: checking its signatures, a standard and its health, verified redaction, and the health copies
/// (optimized, repaired or sanitized) written beside it. Checks and copies leave the document unchanged; a redaction is
/// one undoable change, verified before it is kept.
/// </summary>
internal sealed partial class StudioNativeDocument {
    /// <summary>Checks the document; the request names exactly one of <c>signatures</c>, <c>standard</c> or <c>health</c>.</summary>
    internal async Task<StudioNativeCheckResult> CheckAsync(StudioNativeCheckRequest request, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(request);
        int named = (request.Signatures ? 1 : 0) + (request.Standard is null ? 0 : 1) + (request.Health ? 1 : 0);
        if (named != 1) throw new ArgumentException("Name exactly one check.", nameof(request));
        await EnterAsync(cancellationToken).ConfigureAwait(false);
        try {
            return request switch {
                { Signatures: true } => new StudioNativeCheckResult(StudioNativeAbi.Version, _workspace.Revision,
                    Signatures: Signatures(await _workspace.ValidateSignaturesAsync(cancellationToken).ConfigureAwait(false))),
                { Standard: { } standard } => new StudioNativeCheckResult(StudioNativeAbi.Version, _workspace.Revision,
                    Standard: Standard(await _workspace.AssessComplianceAsync(ProfileOf(standard), cancellationToken).ConfigureAwait(false))),
                _ => new StudioNativeCheckResult(StudioNativeAbi.Version, _workspace.Revision,
                    Health: await Task.Run(() => HealthOf(_workspace.CopyBytes(), _workspace.ReadOptions, cancellationToken), cancellationToken)
                        .ConfigureAwait(false))
            };
        } finally { _gate.Release(); }
    }

    private static StudioNativeSignatureReport Signatures(PdfSignatureValidationReport report) {
        var signatures = report.Signatures.Select(result => {
            PdfSignatureCryptographicResult? crypto = result.CryptographicResult;
            string status = !result.IsStructurallyValid ? "invalid"
                : crypto is null ? "unchecked"
                : crypto.MathematicalSignatureStatus == PdfCryptographicValidationStatus.Invalid || crypto.MessageDigestStatus == PdfCryptographicValidationStatus.Invalid ? "invalid"
                : crypto.MathematicalSignatureStatus == PdfCryptographicValidationStatus.Valid && crypto.MessageDigestStatus == PdfCryptographicValidationStatus.Valid
                    ? (crypto.CertificateChainStatus == PdfCryptographicValidationStatus.Valid ? "valid" : "intact")
                : "unknown";
            // PDF /Name is authored metadata, not the identity from the CMS certificate.
            return new StudioNativeSignature(Subject(crypto?.SignerSubject), result.Signature.FieldName,
                (crypto?.SigningTime)?.ToString("o", CultureInfo.InvariantCulture) ?? result.Signature.SigningTimeRaw,
                result.Signature.Reason, result.Signature.Location, status, result.ByteRangeCoversEndOfFile,
                crypto is null ? null : Status(crypto.CertificateChainStatus), crypto is null ? null : Status(crypto.TimestampStatus),
                result.Findings.Where(static finding => finding.Severity != PdfDiagnosticSeverity.Info).Select(static finding => finding.Message).ToArray());
        }).ToArray();
        string[] findings = report.Findings.Where(static finding => finding.Severity != PdfDiagnosticSeverity.Info)
            .Select(static finding => finding.Message).ToArray();
        return new StudioNativeSignatureReport(report.HasSignatures, signatures, findings);

        static string Status(PdfCryptographicValidationStatus status) => status switch {
            PdfCryptographicValidationStatus.Valid => "valid", PdfCryptographicValidationStatus.Invalid => "invalid",
            PdfCryptographicValidationStatus.NotPerformed => "notChecked", _ => "unknown"
        };

        static string? Subject(string? subject) {
            if (string.IsNullOrWhiteSpace(subject)) return null;
            string? common = subject.Split(',').Select(static part => part.Trim()).FirstOrDefault(static part => part.StartsWith("CN=", StringComparison.OrdinalIgnoreCase));
            return common is null ? subject : common[3..];
        }
    }

    private static PdfComplianceProfile ProfileOf(string standard) => standard switch {
        "pdfa2b" => PdfComplianceProfile.PdfA2B, "pdfa2u" => PdfComplianceProfile.PdfA2U, "pdfa2a" => PdfComplianceProfile.PdfA2A,
        "pdfa3b" => PdfComplianceProfile.PdfA3B, "pdfa3u" => PdfComplianceProfile.PdfA3U, "pdfa3a" => PdfComplianceProfile.PdfA3A,
        "pdfa4" => PdfComplianceProfile.PdfA4, "pdfua1" => PdfComplianceProfile.PdfUa1, "pdfua2" => PdfComplianceProfile.PdfUa2,
        "pdfx1a" => PdfComplianceProfile.PdfX1A2003, "pdfx4" => PdfComplianceProfile.PdfX4,
        _ => throw new ArgumentException($"“{standard}” isn't a standard OfficeIMO checks.", nameof(standard))
    };

    private static StudioNativeStandardReport Standard(PdfComplianceReadinessReport report) =>
        new(report.DisplayName, report.IsReady, report.Requirements.Select(static requirement => new StudioNativeRequirement(requirement.Id,
            requirement.DisplayName, requirement.Status switch {
                PdfComplianceRequirementStatus.Satisfied => "met", PdfComplianceRequirementStatus.Missing => "missing", _ => "unsupported"
            }, requirement.Diagnostic)).ToArray());

    /// <summary>What the document's structure says about it, as the desktop's document health inspection reports it.</summary>
    private static StudioNativeHealth HealthOf(byte[] bytes, PdfLoadOptions options, CancellationToken cancellationToken) {
        PdfDocumentPreflight preflight = OfficeDocument.Inspect(new MemoryStream(bytes), "source.pdf",
            new OfficeInspectOptions { IncludeContent = false, PdfPreflight = options,
                MaximumInputBytes = options.Limits.MaxInputBytes }, cancellationToken).PdfPreflight!;
        PdfDocumentInfo? info = preflight.DocumentInfo;
        var diagnostics = new List<string>(preflight.Diagnostics);
        PdfRepairReport? repairs = null;
        if (preflight.CanRead) {
            try {
                PdfAnalysisReport analysis = PdfDocument.Load(new MemoryStream(bytes), options).Analyze(PdfComplianceProfile.None, cancellationToken);
                info = analysis.Info;
                repairs = analysis.Repair;
                diagnostics.AddRange(repairs.Diagnostics.Select(static item => item.Message));
            } catch (Exception exception) when (exception is not OperationCanceledException and not OutOfMemoryException) {
                diagnostics.Add(exception.Message);
            }
        }
        PdfDocumentSecurityInfo security = info?.Security ?? preflight.Probe.Security;
        return new StudioNativeHealth(bytes.LongLength, info?.PageCount ?? 0, info?.EffectiveVersion ?? info?.HeaderVersion, preflight.CanRead,
            preflight.CanRewrite, security.HasEncryption, security.HasSignatures, info?.HasTaggedContent == true, info?.HasActiveContent == true,
            info?.HasEmbeddedFiles == true, repairs?.RepairCount ?? 0, repairs?.DetectionOnlyCount ?? 0, diagnostics.Distinct().Take(50).ToArray());
    }

    /// <summary>
    /// Writes an optimized, repaired or sanitized copy, verified before it is written: optimization and repair must keep
    /// everything the document contains, and sanitization must remove what it found. Returns what changed.
    /// </summary>
    private async Task<(IReadOnlyList<string> Files, string Summary)> WriteHealthCopyAsync(StudioNativeHealthCopy copy, string kind,
        CancellationToken cancellationToken) {
        if (string.IsNullOrWhiteSpace(copy.Path)) throw new ArgumentException("A destination is required.", nameof(copy));
        byte[] input = _workspace.CopyBytes();
        PdfLoadOptions options = _workspace.ReadOptions;
        (byte[] output, string summary) = await Task.Run(() => {
            switch (kind) {
                case "optimize": {
                    PdfOptimizationOptions optimization = PdfOptimizationOptions.Create(copy.Profile switch {
                        null or "balanced" => PdfOptimizationProfile.Balanced, "smallest" => PdfOptimizationProfile.MaximumCompression,
                        "web" => PdfOptimizationProfile.Web, "archival" => PdfOptimizationProfile.Archival,
                        _ => throw new ArgumentException($"“{copy.Profile}” isn't an optimization profile.", nameof(copy))
                    });
                    optimization.CancellationToken = cancellationToken;
                    PdfOptimizationActionResult result = PdfDocument.Load(new MemoryStream(input), options).Optimization.Apply(optimization);
                    if (!result.PreservationReport.IsPreserved) throw new InvalidOperationException("The optimized copy wouldn't keep everything the document contains, so it wasn't written.");
                    return (result.Bytes, result.ReturnedOriginal ? "The document is already as small as lossless optimization makes it."
                        : string.Create(CultureInfo.InvariantCulture, $"Saved {result.SavedBytes} bytes."));
                }
                case "repair": {
                    PdfRepairArtifactResult result = PdfRepairArtifact.Create(input,
                        new PdfRepairArtifactOptions { CancellationToken = cancellationToken }, options);
                    if (!result.IsVerified) throw new InvalidOperationException("The repaired copy couldn't be verified, so it wasn't written.");
                    return (result.ToBytes(), string.Create(CultureInfo.InvariantCulture, $"Repaired {result.SourceRepairReport.RepairCount} structural problem(s)."));
                }
                default: {
                    PdfSanitizationResult result = PdfDocument.Load(new MemoryStream(input), options).Sanitize(new PdfSanitizationOptions { CancellationToken = cancellationToken });
                    if (!result.IsSanitized || !result.PreservationReport.IsPreserved) throw new InvalidOperationException("The sanitized copy couldn't be verified, so it wasn't written.");
                    return (result.ToBytes(), string.Create(CultureInfo.InvariantCulture,
                        $"Removed {result.RemovedFindings.Count} item(s) such as scripts, actions and attachments; {result.RemainingFindings.Count} remain."));
                }
            }
        }, cancellationToken).ConfigureAwait(false);
        string temporary = copy.Path + ".officeimo-" + Guid.NewGuid().ToString("N")[..8];
        try {
            await File.WriteAllBytesAsync(temporary, output, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            File.Move(temporary, copy.Path, overwrite: true);
        } finally {
            if (File.Exists(temporary)) File.Delete(temporary);
        }
        return ([copy.Path], summary);
    }

    /// <summary>
    /// Finds text to redact, or redacts areas: the request names exactly one of <c>find</c> or <c>apply</c>. Applying
    /// removes the content under each area, paints it over and verifies that the content is gone, as one undoable change.
    /// </summary>
    internal async Task<StudioNativeRedactionResult> RedactAsync(StudioNativeRedactionRequest request, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(request);
        if ((request.Find is null) == (request.Apply is null)) throw new ArgumentException("Name exactly one of find or apply.", nameof(request));
        await EnterAsync(cancellationToken).ConfigureAwait(false);
        try {
            if (request.Find is { } find) {
                IReadOnlyList<PdfRedactionCandidate> candidates = await FindRedactionsAsync(find, cancellationToken).ConfigureAwait(false);
                return new StudioNativeRedactionResult(StudioNativeAbi.Version, _workspace.Revision, candidates
                    .Select(static candidate => new StudioNativeRedactionMark(candidate.Area.PageNumber, StudioNativeRectangle.From(candidate.Bounds), candidate.Description))
                    .ToArray());
            }
            StudioNativeRedactionApply apply = request.Apply!;
            IReadOnlyList<StudioNativeRedactionMark> areas = apply.Areas ?? [];
            if (areas.Count == 0 && apply.Find is null) throw new ArgumentException("Mark at least one area to redact.", nameof(request));
            long expectedRevision = apply.Revision ?? throw new ArgumentException("Include the revision on which the redaction was reviewed.", nameof(request));
            if (expectedRevision != _workspace.Revision)
                throw new InvalidOperationException("The document changed after this redaction was reviewed. Find or mark the content again before applying it.");
            if (!_workspace.CanRedact) throw new NotSupportedException("This document's pages can't be redacted.");
            PdfDocument snapshot = _workspace.CreateDocumentSnapshot();
            var planned = new List<PdfRedactionArea>(areas.Count);
            if (apply.Find is { } search) {
                // Searching again on the same revision keeps the found text's exact glyph areas rather than rectangles around them.
                IReadOnlyList<PdfRedactionCandidate> found = await FindRedactionsAsync(search, cancellationToken).ConfigureAwait(false);
                IReadOnlyList<int> chosen = apply.Include ?? Enumerable.Range(0, found.Count).ToArray();
                foreach (int index in chosen) {
                    if (index < 0 || index >= found.Count) throw new ArgumentOutOfRangeException(nameof(request), "The document changed since the text was found; find it again.");
                    planned.Add(found[index].Area);
                }
            }
            foreach (StudioNativeRedactionMark area in areas) {
                if (area.Page < 1 || area.Page > _workspace.Pages.Count) throw new ArgumentOutOfRangeException(nameof(request), "The page does not exist.");
                StudioNativeRectangle bounds = area.Bounds ?? throw new ArgumentException("Each area needs its bounds.", nameof(request));
                if (!(bounds.Width > 0D && bounds.Height > 0D)) throw new ArgumentException("Each area needs a width and a height.", nameof(request));
                // The editor maps a rectangle on the page as displayed into the page's own space, through its box and rotation.
                var gesture = new PdfEditorGesture(area.Page, Finite(bounds.X), Finite(bounds.Y), Finite(bounds.X + bounds.Width), Finite(bounds.Y + bounds.Height), []);
                PdfEditorCommand command = PdfEditorCommandFactory.Create(snapshot, PdfEditorTool.Redact, gesture,
                    new PdfEditorProperties(string.Empty, string.Empty, global::OfficeIMO.OfficeColor.Black, string.Empty, string.Empty, 12D));
                planned.Add(new PdfRedactionArea(area.Page, command.Bounds.Left, command.Bounds.Bottom, command.Bounds.Width, command.Bounds.Height,
                    "OfficeIMO Studio redaction"));
            }
            if (planned.Count == 0) throw new ArgumentException("Mark at least one area to redact.", nameof(request));
            PdfRedactionPlan plan = await _workspace.PlanRedactionsAsync(planned, cancellationToken).ConfigureAwait(false);
            PdfSanitizationOptions? sanitization = apply.Sanitize ? new PdfSanitizationOptions {
                ContentKindsToRemove = PdfSanitizationContentKind.UserMetadata | PdfSanitizationContentKind.EmbeddedFiles |
                    PdfSanitizationContentKind.Actions | PdfSanitizationContentKind.CommentsAndMarkup
            } : null;
            PdfVerifiedRedactionResult verified = await _workspace.ApplyVerifiedRedactionAsync(plan, expectedRevision,
                string.IsNullOrWhiteSpace(apply.RemovedText) ? null : apply.RemovedText, cancellationToken, sanitizationOptions: sanitization).ConfigureAwait(false);
            PdfRedactionEvidenceReport evidence = verified.Evidence;
            return new StudioNativeRedactionResult(StudioNativeAbi.Version, _workspace.Revision, Verification: new StudioNativeRedactionVerification(
                evidence.IsVerified, evidence.VerifiedAbsentCount, evidence.ResidualCount, evidence.InconclusiveCount, evidence.AffectedPageNumbers.ToArray()));
        } finally { _gate.Release(); }
    }

    // Matches the found text's own glyphs, as Preview's redaction does, unless the user chose whole blocks because the
    // glyphs can't be removed without touching their neighbours.
    private Task<IReadOnlyList<PdfRedactionCandidate>> FindRedactionsAsync(StudioNativeRedactionFind find, CancellationToken cancellationToken) {
        if (string.IsNullOrWhiteSpace(find.Text)) throw new ArgumentException("Enter the text to find.", nameof(find));
        IReadOnlyCollection<int>? pages = find.Pages is { Count: > 0 } chosen ? RequirePages(chosen) : null;
        return _workspace.SearchRedactionMarksAsync(find.Text, find.Regex, find.MatchCase, pages, cancellationToken,
            find.WholeBlocks ? PdfRedactionTextSelection.LogicalBlocks : PdfRedactionTextSelection.MatchedGlyphs);
    }
}

internal sealed record StudioNativeCheckRequest(bool Signatures = false, string? Standard = null, bool Health = false);

internal sealed record StudioNativeSignature(string? Signer, string? Field, string? SigningTime, string? Reason, string? Location, string Status,
    bool CoversWholeDocument, string? Certificate, string? Timestamp, IReadOnlyList<string> Findings);

internal sealed record StudioNativeSignatureReport(bool HasSignatures, IReadOnlyList<StudioNativeSignature> Signatures, IReadOnlyList<string> Findings);

internal sealed record StudioNativeRequirement(string Id, string Name, string Status, string Detail);

internal sealed record StudioNativeStandardReport(string Name, bool IsReady, IReadOnlyList<StudioNativeRequirement> Requirements);

internal sealed record StudioNativeHealth(long FileSize, int Pages, string? Version, bool CanRead, bool CanRewrite, bool IsEncrypted, bool IsSigned,
    bool IsTagged, bool HasActiveContent, bool HasAttachments, int RecoveredProblems, int UnrecoverableProblems, IReadOnlyList<string> Diagnostics);

internal sealed record StudioNativeCheckResult(int AbiVersion, long Revision, StudioNativeSignatureReport? Signatures = null,
    StudioNativeStandardReport? Standard = null, StudioNativeHealth? Health = null);

/// <summary>A health copy at <c>path</c>; <c>profile</c> chooses an optimization: balanced (default), smallest, web or archival.</summary>
internal sealed record StudioNativeHealthCopy(string? Path, string? Profile = null);

internal sealed record StudioNativeRedactionFind(string? Text, bool Regex = false, bool MatchCase = false, IReadOnlyList<int>? Pages = null, bool WholeBlocks = false);

/// <summary>An area to redact, in points of the page as displayed; <c>preview</c> is the text a search found there.</summary>
internal sealed record StudioNativeRedactionMark(int Page, StudioNativeRectangle? Bounds, string? Preview = null);

/// <summary>Areas or search matches reviewed on the required <c>revision</c>; a stale selection is rejected before planning.</summary>
internal sealed record StudioNativeRedactionApply(IReadOnlyList<StudioNativeRedactionMark>? Areas = null, StudioNativeRedactionFind? Find = null,
    IReadOnlyList<int>? Include = null, string? RemovedText = null, bool Sanitize = false, long? Revision = null);

internal sealed record StudioNativeRedactionRequest(StudioNativeRedactionFind? Find = null, StudioNativeRedactionApply? Apply = null);

internal sealed record StudioNativeRedactionVerification(bool Verified, int Removed, int Residual, int Inconclusive, IReadOnlyList<int> Pages);

internal sealed record StudioNativeRedactionResult(int AbiVersion, long Revision, IReadOnlyList<StudioNativeRedactionMark>? Marks = null,
    StudioNativeRedactionVerification? Verification = null);
