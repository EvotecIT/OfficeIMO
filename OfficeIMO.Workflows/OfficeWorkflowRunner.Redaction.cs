using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner : IPdfRedactionWorkflowRunner {
    /// <inheritdoc />
    public async Task<PdfRedactionWorkflowResult> RunRedactionAsync(
        PdfRedactionWorkflowRequest request,
        IProgress<OfficeWorkflowProgress>? progress = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(request);
        PdfRedactionWorkflowRequest executionRequest = SnapshotRedactionRequest(request);
        PreparedRedactionResult prepared;
        try {
            prepared = await PrepareRedactionAsync(executionRequest, progress, cancellationToken).ConfigureAwait(false);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            return FailedRedactionResult(executionRequest, OfficeWorkflowStatus.Cancelled, "Redaction workflow cancelled.", "Cancelled", OfficeWorkflowDiagnosticSeverity.Information);
        } catch (Exception exception) when (exception is not OutOfMemoryException and not StackOverflowException) {
            return FailedRedactionResult(executionRequest, OfficeWorkflowStatus.Failed, "Redaction workflow failed.", "RedactionWorkflowFailed", OfficeWorkflowDiagnosticSeverity.Error, exception);
        }

        try {
            cancellationToken.ThrowIfCancellationRequested();
            PublishPreparedFiles(prepared.Files, executionRequest.ConflictPolicy, cancellationToken);
            Report(progress, executionRequest.Id, "complete", prepared.Result.Summary, 1D);
            return prepared.Result;
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            return FailedRedactionResult(executionRequest, OfficeWorkflowStatus.Cancelled, "Redaction workflow cancelled before publication.", "Cancelled", OfficeWorkflowDiagnosticSeverity.Information);
        } catch (RedactionBackupCleanupException exception) {
            return MarkPublishedWithCleanupFailure(prepared.Result, exception);
        } catch (Exception exception) when (exception is not OutOfMemoryException and not StackOverflowException) {
            return FailedRedactionResult(executionRequest, OfficeWorkflowStatus.Failed, "Redaction publication failed.", "RedactionPublicationFailed", OfficeWorkflowDiagnosticSeverity.Error, exception);
        }
    }

    /// <inheritdoc />
    public async Task<PdfRedactionBatchResult> RunRedactionBatchAsync(
        IEnumerable<PdfRedactionWorkflowRequest> requests,
        IProgress<OfficeWorkflowProgress>? progress = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(requests);
        const int absoluteMaximumBatchItems = 10_000;
        var snapshots = new List<PdfRedactionWorkflowRequest>();
        int configuredMaximum = absoluteMaximumBatchItems;
        foreach (PdfRedactionWorkflowRequest supplied in requests) {
            cancellationToken.ThrowIfCancellationRequested();
            if (supplied is null) throw new ArgumentException("Batch requests cannot contain null entries.", nameof(requests));
            PdfRedactionWorkflowRequest snapshot = SnapshotRedactionRequest(supplied);
            configuredMaximum = Math.Min(configuredMaximum, snapshot.Limits.MaximumBatchItems);
            if (snapshots.Count >= configuredMaximum || snapshots.Count >= absoluteMaximumBatchItems) {
                throw new RedactionWorkflowException($"The redaction batch exceeds the configured {configuredMaximum}-item limit.");
            }
            snapshots.Add(snapshot);
        }
        PdfRedactionWorkflowRequest[] batch = snapshots.ToArray();
        if (batch.Any(static request => request.Limits is null || request.Limits.MaximumBatchItems <= 0 || request.Limits.MaximumConcurrency <= 0 || request.Limits.MaximumConcurrency > 32 || request.Limits.MaximumBatchPreparedBytes <= 0)) throw new ArgumentException("Every batch request requires valid positive batch limits and MaximumConcurrency from 1 through 32.", nameof(requests));
        if (batch.Length == 0) return new PdfRedactionBatchResult(OfficeWorkflowStatus.Completed, Array.Empty<PdfRedactionWorkflowResult>(), true, "The redaction batch was empty.");
        OfficeWorkflowConflictPolicy batchConflictPolicy = batch[0].ConflictPolicy;
        if (batchConflictPolicy == OfficeWorkflowConflictPolicy.Rename || batch.Any(request => request.ConflictPolicy != batchConflictPolicy)) throw new ArgumentException("An atomic redaction batch requires one shared Fail or Replace conflict policy.", nameof(requests));

        int maximumConcurrency = batch.Min(static request => request.Limits.MaximumConcurrency);
        long maximumPreparedBytes = batch.Min(static request => request.Limits.MaximumBatchPreparedBytes);
        var preparedBudget = new PreparedByteBudget(maximumPreparedBytes);
        var prepared = new PreparedRedactionResult?[batch.Length];
        var reservations = new PreparedByteReservation?[batch.Length];
        var errors = new Exception?[batch.Length];
        using (var gate = new SemaphoreSlim(maximumConcurrency, maximumConcurrency)) {
            Task[] preparations = batch.Select((request, index) => Task.Run(async () => {
                try {
                    await gate.WaitAsync(cancellationToken).ConfigureAwait(false);
                    try {
                        long maximumItemBytes = GetMaximumPreparedBytes(request);
                        using PreparedByteReservation pendingReservation = preparedBudget.Reserve(maximumItemBytes);
                        int batchIndex = index;
                        var itemProgress = progress is null ? null : new InlineProgress<OfficeWorkflowProgress>(item => progress.Report(
                            new OfficeWorkflowProgress(item.RequestId, item.Stage, $"{batchIndex + 1} of {batch.Length} · {item.Message}", item.Fraction, (batchIndex + item.Fraction) / batch.Length)));
                        PreparedRedactionResult item = await PrepareRedactionAsync(request, itemProgress, cancellationToken).ConfigureAwait(false);
                        pendingReservation.Resize(item.Files.Sum(static file => (long)file.Bytes.Length));
                        reservations[index] = pendingReservation.Transfer();
                        prepared[index] = item;
                    } finally {
                        gate.Release();
                    }
                } catch (Exception exception) when (exception is not OutOfMemoryException and not StackOverflowException) {
                    errors[index] = exception;
                }
            }, CancellationToken.None)).ToArray();
            await Task.WhenAll(preparations).ConfigureAwait(false);
        }
        if (errors.Any(static error => error is not null)) {
            PdfRedactionWorkflowResult[] results = batch.Select((request, index) => prepared[index]?.Result ??
                (errors[index] is OperationCanceledException && cancellationToken.IsCancellationRequested
                    ? FailedRedactionResult(request, OfficeWorkflowStatus.Cancelled, "Redaction batch cancelled before publication.", "Cancelled", OfficeWorkflowDiagnosticSeverity.Information)
                    : FailedRedactionResult(request, OfficeWorkflowStatus.Failed, "Redaction batch preparation failed.", "RedactionWorkflowFailed", OfficeWorkflowDiagnosticSeverity.Error, errors[index]))).ToArray();
            OfficeWorkflowStatus status = cancellationToken.IsCancellationRequested ? OfficeWorkflowStatus.Cancelled : OfficeWorkflowStatus.Failed;
            for (int index = 0; index < results.Length; index++) {
                if (prepared[index] is not null) results[index] = MarkNotPublished(prepared[index]!.Result, status, "No artifact was published because the atomic batch did not complete.");
            }
            DisposeReservations(reservations);
            return new PdfRedactionBatchResult(status, results, false, "No batch artifacts were published because one or more items did not prepare successfully.");
        }
        PreparedRedactionResult[] completedPreparations = prepared.Select(static item => item!).ToArray();

        try {
            List<PreparedFile> files = completedPreparations.SelectMany(static item => item.Files).ToList();
            EnsureUniqueDestinations(files);
            EnsureDestinationsDoNotReplaceReviewedInputs(files, batch);
            PublishPreparedFiles(files, batchConflictPolicy, cancellationToken);
            DisposeReservations(reservations);
            return new PdfRedactionBatchResult(OfficeWorkflowStatus.Completed, completedPreparations.Select(static item => item.Result).ToArray(), true, $"Published {batch.Length} redaction workflow item(s) atomically.");
        } catch (RedactionBackupCleanupException exception) {
            DisposeReservations(reservations);
            PdfRedactionWorkflowResult[] affected = completedPreparations.Select(item => MarkPublishedWithCleanupFailure(item.Result, exception)).ToArray();
            return new PdfRedactionBatchResult(OfficeWorkflowStatus.Failed, affected, true, "The atomic batch was published, but prior-destination rollback data could not be removed. Host cleanup is required.");
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            DisposeReservations(reservations);
            return new PdfRedactionBatchResult(OfficeWorkflowStatus.Cancelled, completedPreparations.Select(static item => MarkNotPublished(item.Result, OfficeWorkflowStatus.Cancelled, "The atomic batch was cancelled; no artifact was published.")).ToArray(), false, "The batch was cancelled and publication was rolled back.");
        } catch (Exception exception) when (exception is not OutOfMemoryException and not StackOverflowException) {
            DisposeReservations(reservations);
            PdfRedactionWorkflowResult[] failed = completedPreparations.Select(item => MarkNotPublished(item.Result, OfficeWorkflowStatus.Failed, "Atomic batch publication failed; no artifact was committed.", exception)).ToArray();
            return new PdfRedactionBatchResult(OfficeWorkflowStatus.Failed, failed, false, "Atomic batch publication failed and published files were rolled back.");
        }
    }

    private static async Task<PreparedRedactionResult> PrepareRedactionAsync(
        PdfRedactionWorkflowRequest request,
        IProgress<OfficeWorkflowProgress>? progress,
        CancellationToken cancellationToken) {
        ValidateRedactionRequest(request);
        Report(progress, request.Id, "validate", "Validating source, recipe, and resource limits", 0.05D);
        byte[] originalSource = await ReadFileBoundedAsync(request.InputPath, request.Limits.MaximumInputBytes, cancellationToken).ConfigureAwait(false);
        string originalSourceSha = ComputeSha256(originalSource);
        string recipeSha = ComputeRecipeSha256(request.Recipe);
        PdfLoadOptions loadOptions = new() { Password = request.OwnerPassword };
        PdfDocumentPreflight originalPreflight = PdfInspector.Preflight(originalSource, loadOptions);
        int sourceSignatureCount = Math.Max(
            originalPreflight.Probe.Security.SignatureCount,
            originalPreflight.Probe.HasSignatures ? 1 : 0);
        if (originalPreflight.Probe.HasSignatures && request.Recipe.SignaturePolicy == PdfRedactionSignaturePolicy.RejectSignedSource) {
            throw new RedactionWorkflowException("The recipe rejects signed sources because permanent redaction invalidates their signatures. Select an explicit derivative policy.");
        }
        if (request.Mode == PdfRedactionWorkflowMode.ApplyAndVerify &&
            request.Recipe.SignaturePolicy == PdfRedactionSignaturePolicy.CreateAndSignDerivative &&
            request.OutputSigner is null) {
            throw new RedactionWorkflowException("CreateAndSignDerivative requires a runtime output signer.");
        }

        byte[] planningSource = originalSource;
        PdfLoadOptions planningOptions = loadOptions;
        if (originalPreflight.Probe.HasEncryption) {
            if (request.Recipe.EncryptedDocumentPolicy == PdfRedactionEncryptedDocumentPolicy.Reject) {
                throw new RedactionWorkflowException("The recipe rejects encrypted PDFs. Select an explicit authenticated encryption policy and supply the owner password at runtime.");
            }
            if (string.IsNullOrWhiteSpace(request.OwnerPassword)) throw new RedactionWorkflowException("An owner password is required by the encrypted-document policy.");
            if (!originalPreflight.Probe.Security.HasOwnerAuthorization) throw new RedactionWorkflowException("The supplied credential did not authenticate as the PDF owner password.");
            if (originalPreflight.Probe.HasSignatures) {
                planningSource = PdfDocument.Load(originalSource, loadOptions).Security.CreateUnsignedDerivative(cancellationToken).Pdf;
                planningOptions = PdfLoadOptions.Default;
            } else if (request.Recipe.EncryptedDocumentPolicy is PdfRedactionEncryptedDocumentPolicy.Decrypt or PdfRedactionEncryptedDocumentPolicy.DecryptAndReencrypt) {
                planningSource = PdfDocument.Load(originalSource, loadOptions).Security.Decrypt(request.OwnerPassword).Pdf;
                planningOptions = PdfLoadOptions.Default;
            }
        } else if (request.Recipe.EncryptedDocumentPolicy != PdfRedactionEncryptedDocumentPolicy.Reject) {
            throw new RedactionWorkflowException("Decrypt and DecryptAndReencrypt policies require an encrypted source PDF.");
        }
        if (originalPreflight.Probe.HasSignatures && !originalPreflight.Probe.HasEncryption) {
            planningSource = PdfDocument.Load(originalSource, loadOptions).Security.CreateUnsignedDerivative(cancellationToken).Pdf;
            planningOptions = PdfLoadOptions.Default;
        }
        if (request.Mode != PdfRedactionWorkflowMode.PlanOnly && request.Recipe.EncryptedDocumentPolicy == PdfRedactionEncryptedDocumentPolicy.DecryptAndReencrypt && request.OutputEncryption is null) {
            throw new RedactionWorkflowException("DecryptAndReencrypt apply and verification require runtime OutputEncryption settings.");
        }

        cancellationToken.ThrowIfCancellationRequested();
        Report(progress, request.Id, "plan", "Building source-bound native and region candidates", 0.2D);
        PdfDocument document = PdfDocument.Load(planningSource, planningOptions);
        CandidatePlan candidatePlan = await BuildCandidatePlanAsync(document, request, originalSourceSha, cancellationToken).ConfigureAwait(false);
        if (candidatePlan.Candidates.Count > request.Limits.MaximumCandidates) throw new RedactionWorkflowException($"The plan produced {candidatePlan.Candidates.Count} candidates, above the configured {request.Limits.MaximumCandidates}-candidate limit.");
        if (candidatePlan.AreaCount > request.Limits.MaximumAreas) throw new RedactionWorkflowException($"The plan produced {candidatePlan.AreaCount} areas, above the configured {request.Limits.MaximumAreas}-area limit.");

        if (request.Mode == PdfRedactionWorkflowMode.PlanOnly) {
            var result = new PdfRedactionWorkflowResult(request.Id, request.Mode, OfficeWorkflowStatus.Completed, $"Planned {candidatePlan.Candidates.Count} privacy-safe redaction candidate(s).", originalSourceSha, recipeSha, candidatePlan.Candidates, null, NormalizeOptionalPath(request.EvidencePath), null, candidatePlan.Diagnostics);
            return CreatePreparedResult(result, request.EvidencePath, request.Limits, outputBytes: null);
        }

        ValidateDecisions(request.Decisions!, originalSourceSha, recipeSha, candidatePlan.Candidates);
        HashSet<string> approvedIds = new(request.Decisions!.ApprovedCandidateIds, StringComparer.Ordinal);
        PdfRedactionArea[] approvedAreas = candidatePlan.CandidateAreas
            .Where((_, index) => approvedIds.Contains(candidatePlan.Candidates[index].Id))
            .SelectMany(static areas => areas)
            .ToArray();
        Report(progress, request.Id, request.Mode == PdfRedactionWorkflowMode.ApplyAndVerify ? "apply" : "verify", "Using the reviewed source-bound candidate set", 0.48D);

        byte[] output;
        PdfRedactionWorkflowEvidence evidence;
        if (request.Mode == PdfRedactionWorkflowMode.ApplyAndVerify) {
            if (approvedAreas.Length == 0) {
                output = planningSource.ToArray();
                if (request.Recipe.EncryptedDocumentPolicy == PdfRedactionEncryptedDocumentPolicy.DecryptAndReencrypt) {
                    output = PdfDocument.Load(output).Security.Encrypt(request.OutputEncryption!).Pdf;
                }
                if (output.LongLength > request.Limits.MaximumOutputBytes) throw new RedactionWorkflowException("Redacted output exceeds the configured output-byte limit.");
                PdfLoadOptions finalOutputOptions = GetOutputLoadOptions(request);
                ValidateFinalArtifactEncryptionPolicy(output, request, finalOutputOptions, cancellationToken);
                (output, RedactionSignatureEvidence signature) = ApplyDerivativeSignature(output, request, finalOutputOptions, sourceSignatureCount, cancellationToken);
                if (output.LongLength > request.Limits.MaximumOutputBytes) throw new RedactionWorkflowException("Redacted output exceeds the configured output-byte limit after signing.");
                ValidateFinalArtifactEncryptionPolicy(output, request, finalOutputOptions, cancellationToken);
                IReadOnlyList<string> externalValidators = ValidateExternalArtifact(output, request, finalOutputOptions, cancellationToken);
                evidence = CreateEmptyEvidence(originalSourceSha, ComputeSha256(output), recipeSha, request, candidatePlan,
                    originalPreflight.Probe.HasEncryption ? request.Recipe.EncryptedDocumentPolicy.ToString() : "NoRewrite", signature, externalValidators);
            } else {
                PdfRedactionPlan approvedPlan = document.Redactions.Plan(approvedAreas, cancellationToken);
                var applyOptions = new PdfRedactionApplyOptions {
                    CancellationToken = cancellationToken,
                    CleanupScope = request.Recipe.CleanupScope,
                    RemoveIntersectingPaths = request.Recipe.RemoveIntersectingPaths,
                    UnsupportedImagePolicy = request.Recipe.UnsupportedImagePolicy
                };
                PdfRedactionApplyResult applied = document.Redactions.ApplyWithEvidence(
                    approvedPlan,
                    applyOptions,
                    new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true, CheckManagedRendering = true, CancellationToken = cancellationToken });
                cancellationToken.ThrowIfCancellationRequested();
                output = applied.Pdf;
                if (request.Recipe.EncryptedDocumentPolicy == PdfRedactionEncryptedDocumentPolicy.DecryptAndReencrypt) {
                    output = PdfDocument.Load(output).Security.Encrypt(request.OutputEncryption!).Pdf;
                }
                PdfLoadOptions finalOutputOptions = GetOutputLoadOptions(request);
                ValidateFinalArtifactEncryptionPolicy(output, request, finalOutputOptions, cancellationToken);
                PdfRedactionVerificationReport finalVerification = PdfDocument.Load(output, finalOutputOptions).Redactions.VerifyAppliedPlan(
                    approvedPlan,
                    new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true, CheckManagedRendering = true, CancellationToken = cancellationToken });
                if (!finalVerification.IsVerified) throw new RedactionWorkflowException("The final protected-document artifact failed redaction verification; no artifact will be published.");
                (output, RedactionSignatureEvidence signature) = ApplyDerivativeSignature(output, request, finalOutputOptions, sourceSignatureCount, cancellationToken);
                ValidateFinalArtifactEncryptionPolicy(output, request, finalOutputOptions, cancellationToken);
                IReadOnlyList<string> externalValidators = ValidateExternalArtifact(output, request, finalOutputOptions, cancellationToken);
                int ocrResidualCount = candidatePlan.OcrUsed
                    ? await CountOcrResidualsAsync(PdfDocument.Load(output, finalOutputOptions), request, approvedAreas, cancellationToken).ConfigureAwait(false)
                    : 0;
                if (output.LongLength > request.Limits.MaximumOutputBytes) throw new RedactionWorkflowException("Redacted output exceeds the configured output-byte limit.");
                evidence = CreateEvidence(originalSourceSha, ComputeSha256(output), recipeSha, request, candidatePlan, applied.Evidence, ocrResidualCount, signature, externalValidators);
                if (!evidence.Verified) throw new RedactionWorkflowException("Redaction verification was inconclusive or found residual content; no artifact will be published.");
            }
        } else {
            byte[] existingOutput = await ReadFileBoundedAsync(request.OutputPath!, request.Limits.MaximumOutputBytes, cancellationToken).ConfigureAwait(false);
            PdfLoadOptions outputOptions = GetOutputLoadOptions(request);
            ValidateFinalArtifactEncryptionPolicy(existingOutput, request, outputOptions, cancellationToken);
            RedactionSignatureEvidence signature = InspectDerivativeSignature(existingOutput, outputOptions, request, sourceSignatureCount);
            IReadOnlyList<string> externalValidators = ValidateExternalArtifact(existingOutput, request, outputOptions, cancellationToken);
            (byte[] verificationOutput, PdfLoadOptions verificationOptions) = CreateSignatureFreeVerificationArtifact(existingOutput, outputOptions, request, cancellationToken);
            if (approvedAreas.Length == 0) {
                string existingOutputSha = ComputeSha256(existingOutput);
                if (request.Recipe.EncryptedDocumentPolicy == PdfRedactionEncryptedDocumentPolicy.DecryptAndReencrypt ||
                    request.Recipe.SignaturePolicy == PdfRedactionSignaturePolicy.CreateAndSignDerivative) {
                    if (string.IsNullOrWhiteSpace(request.ExpectedOutputSha256) || !string.Equals(request.ExpectedOutputSha256, existingOutputSha, StringComparison.OrdinalIgnoreCase)) {
                        throw new RedactionWorkflowException("Zero-area verification of a re-encrypted artifact requires its trusted expected output SHA-256 from prior apply evidence.");
                    }
                } else {
                    byte[] expected = originalPreflight.Probe.HasEncryption ? planningSource : originalSource;
                    if (!expected.AsSpan().SequenceEqual(verificationOutput)) throw new RedactionWorkflowException("With no approved candidates, existing-output verification requires the exact policy-transformed source artifact.");
                }
                evidence = CreateEmptyEvidence(originalSourceSha, existingOutputSha, recipeSha, request, candidatePlan,
                    originalPreflight.Probe.HasEncryption ? request.Recipe.EncryptedDocumentPolicy.ToString() : "NoRewrite", signature, externalValidators);
            } else {
                PdfRedactionPlan approvedPlan = document.Redactions.Plan(approvedAreas, cancellationToken);
                PdfRedactionVerificationReport verification = PdfDocument.Load(verificationOutput, verificationOptions).Redactions.VerifyAppliedPlan(
                    approvedPlan,
                    new PdfRedactionVerificationOptions { RequireCompleteStreamInspection = true, CheckManagedRendering = true, CancellationToken = cancellationToken });
                int ocrResidualCount = candidatePlan.OcrUsed
                    ? await CountOcrResidualsAsync(PdfDocument.Load(verificationOutput, verificationOptions), request, approvedAreas, cancellationToken).ConfigureAwait(false)
                    : 0;
                evidence = CreateVerificationEvidence(originalSourceSha, ComputeSha256(existingOutput), recipeSha, request, candidatePlan, approvedPlan, verification, ocrResidualCount, signature, externalValidators);
            }
            output = Array.Empty<byte>();
            if (!evidence.Verified) throw new RedactionWorkflowException("Existing-output verification was inconclusive or found residual content.");
        }

        string? publishedOutput = request.Mode == PdfRedactionWorkflowMode.ApplyAndVerify ? NormalizeOptionalPath(request.OutputPath) : NormalizeOptionalPath(request.OutputPath);
        string? evidencePath = NormalizeOptionalPath(request.EvidencePath);
        var completed = new PdfRedactionWorkflowResult(request.Id, request.Mode, OfficeWorkflowStatus.Completed,
            request.Mode == PdfRedactionWorkflowMode.ApplyAndVerify ? $"Applied and verified {approvedAreas.Length} approved candidate(s)." : $"Verified {approvedAreas.Length} approved candidate(s) against the existing output.",
            originalSourceSha, recipeSha, candidatePlan.Candidates, publishedOutput, evidencePath, evidence, candidatePlan.Diagnostics);
        return CreatePreparedResult(completed, request.EvidencePath, request.Limits, request.Mode == PdfRedactionWorkflowMode.ApplyAndVerify ? output : null, request.OutputPath);
    }

    private static async Task<CandidatePlan> BuildCandidatePlanAsync(PdfDocument document, PdfRedactionWorkflowRequest request, string sourceSha, CancellationToken cancellationToken) {
        PdfRedactionSearchOptions search = BuildSearchOptions(request.Recipe, request.Limits.MaximumCandidates, cancellationToken);
        var candidateAreaSets = new List<IReadOnlyList<PdfRedactionArea>>();
        var origins = new List<CandidateOrigin>();
        var diagnostics = new List<OfficeWorkflowDiagnostic>();
        bool hasSearch = search.LiteralText.Count > 0 || search.RegularExpressions.Count > 0 || search.FormFieldNames.Count > 0 || search.LogicalElementKinds.Count > 0;
        bool runNative = hasSearch && request.Recipe.DetectionMode != PdfRedactionDetectionMode.OcrOnly;
        int nativeCount = 0;
        if (runNative) {
            PdfRedactionPlan nativePlan = document.Redactions.Search(search);
            foreach (PdfRedactionArea area in nativePlan.Areas) { candidateAreaSets.Add(new[] { area }); origins.Add(new CandidateOrigin("native", null, null, null, null)); }
            nativeCount = nativePlan.Areas.Count;
        }

        foreach (PdfRedactionRecipeRegion region in request.Recipe.Regions) {
            PdfRedactionRegion normalized = ConvertRegion(region);
            candidateAreaSets.Add(normalized.Areas);
            origins.Add(new CandidateOrigin("region:" + normalized.Kind, null, null, null, null));
        }

        bool runOcr = request.Recipe.DetectionMode == PdfRedactionDetectionMode.OcrOnly ||
            request.Recipe.DetectionMode == PdfRedactionDetectionMode.NativeAndOcr ||
            request.Recipe.DetectionMode == PdfRedactionDetectionMode.NativeThenOcr && nativeCount == 0;
        if (runOcr) {
            if (request.OcrEngine is null) throw new RedactionWorkflowException("The recipe requires OCR but no runtime OCR engine was supplied.");
            PdfOcrRedactionSearchResult ocr = await document.SearchRedactionCandidatesWithOcrAsync(request.OcrEngine, search, request.OcrOptions, cancellationToken).ConfigureAwait(false);
            foreach (PdfOcrRedactionCandidate candidate in ocr.Candidates) {
                candidateAreaSets.Add(new[] { candidate.Area });
                origins.Add(new CandidateOrigin("ocr", candidate.MinimumConfidence, candidate.Provider, candidate.Model, candidate.Language));
            }
            diagnostics.Add(new OfficeWorkflowDiagnostic("OcrCandidateDiscovery", $"OCR contributed {ocr.Candidates.Count} privacy-safe candidate(s).", details: new Dictionary<string, string> { ["provider"] = request.OcrEngine.Id, ["candidateCount"] = ocr.Candidates.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) }));
        }

        if (request.Recipe.Rules.Any(static rule => rule.Kind == PdfRedactionRuleKind.RedactAnnotations)) {
            foreach (PdfAnnotation annotation in document.Reader.AnnotationsBySubtype("Redact")
                .Where(static annotation => annotation.PageNumber.HasValue && (annotation.QuadPoints.Count >= 8 || annotation.Width > 0D && annotation.Height > 0D))) {
                candidateAreaSets.Add(PdfRedactionRegion.FromRedactAnnotation(annotation).Areas);
                origins.Add(new CandidateOrigin("annotation", null, null, null, null));
            }
        }

        var candidateAreas = new List<IReadOnlyList<PdfRedactionArea>>();
        var candidates = new List<PdfRedactionWorkflowCandidate>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        for (int index = 0; index < candidateAreaSets.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            IReadOnlyList<PdfRedactionArea> areas = candidateAreaSets[index];
            CandidateOrigin origin = origins[index];
            string originIdentity = string.Join("|", origin.Origin, origin.Provider ?? string.Empty, origin.Model ?? string.Empty, origin.Language ?? string.Empty, origin.Confidence?.ToString("R", System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty);
            string id = ComputeCandidateId(sourceSha, originIdentity, areas);
            if (!seen.Add(id)) continue;
            candidateAreas.Add(areas);
            candidates.Add(new PdfRedactionWorkflowCandidate(id, origin.Origin, areas, origin.Confidence, origin.Provider, origin.Model, origin.Language));
        }
        return new CandidatePlan(candidateAreas, candidates, diagnostics, runOcr);
    }

    private static PdfRedactionSearchOptions BuildSearchOptions(PdfRedactionRecipe recipe, int maximumCandidates, CancellationToken cancellationToken) {
        var search = new PdfRedactionSearchOptions { MatchCase = recipe.MatchCase, RegexTimeout = TimeSpan.FromMilliseconds(recipe.RegexTimeoutMilliseconds), MaximumCandidates = maximumCandidates, CancellationToken = cancellationToken };
        foreach (PdfRedactionRule rule in recipe.Rules) {
            switch (rule.Kind) {
                case PdfRedactionRuleKind.Literal:
                    search.AddLiteral(rule.Value!);
                    break;
                case PdfRedactionRuleKind.Regex:
                    search.AddRegex(rule.Value!);
                    break;
                case PdfRedactionRuleKind.FormField:
                    search.AddFormField(rule.Value!);
                    break;
                case PdfRedactionRuleKind.LogicalKind:
                    if (!Enum.TryParse(rule.Value!, ignoreCase: true, out PdfLogicalElementKind kind) || !Enum.IsDefined(kind)) throw new ArgumentException("Unknown logical element kind in redaction recipe.");
                    search.AddLogicalKind(kind);
                    break;
                case PdfRedactionRuleKind.RedactAnnotations:
                    break;
                default:
                    throw new ArgumentException("Unknown redaction rule kind.");
            }
        }
        return search;
    }

    private static PdfRedactionRegion ConvertRegion(PdfRedactionRecipeRegion region) {
        if (region is null) throw new ArgumentException("Recipe regions cannot contain null entries.");
        PdfRedactionPoint[] points = region.Points.Select(static point => new PdfRedactionPoint(point.X, point.Y)).ToArray();
        return region.Kind switch {
            PdfRedactionRegionKind.Rectangle => PdfRedactionRegion.Rectangle(region.PageNumber, region.X, region.Y, region.Width, region.Height, region.Label),
            PdfRedactionRegionKind.Quadrilateral => PdfRedactionRegion.Quadrilateral(region.PageNumber, points, region.Label),
            PdfRedactionRegionKind.Polygon => PdfRedactionRegion.Polygon(region.PageNumber, points, region.Label),
            PdfRedactionRegionKind.Freehand => PdfRedactionRegion.Freehand(region.PageNumber, points, region.StrokeWidth, region.Label),
            PdfRedactionRegionKind.Group => PdfRedactionRegion.Group(region.PageNumber, region.Areas.Select(ConvertRegion).SelectMany(static item => item.Areas), region.Label),
            _ => throw new ArgumentException("Unknown redaction region kind.")
        };
    }

    private static void ValidateRedactionRequest(PdfRedactionWorkflowRequest request) {
        ArgumentNullException.ThrowIfNull(request.Recipe);
        ArgumentNullException.ThrowIfNull(request.Limits);
        ArgumentNullException.ThrowIfNull(request.ExternalValidators);
        if (!Enum.IsDefined(request.Mode)) throw new ArgumentException("Redaction workflow mode is not defined.");
        if (!Enum.IsDefined(request.ConflictPolicy)) throw new ArgumentException("Redaction conflict policy is not defined.");
        if (!Enum.IsDefined(request.Recipe.DetectionMode) ||
            (request.Recipe.CleanupScope & ~PdfRedactionCleanupScope.All) != 0 ||
            !Enum.IsDefined(request.Recipe.UnsupportedImagePolicy) || !Enum.IsDefined(request.Recipe.EncryptedDocumentPolicy) ||
            !Enum.IsDefined(request.Recipe.SignaturePolicy)) {
            throw new ArgumentException("The redaction recipe contains an undefined policy value.");
        }
        if (request.OutputEncryption is not null && !Enum.IsDefined(request.OutputEncryption.Algorithm)) throw new ArgumentException("Output encryption algorithm is not defined.");
        if (string.IsNullOrWhiteSpace(request.Id)) throw new ArgumentException("Request id cannot be empty.");
        if (string.IsNullOrWhiteSpace(request.InputPath) || !File.Exists(request.InputPath)) throw new FileNotFoundException("Redaction input PDF was not found.", request.InputPath);
        if (!string.Equals(Path.GetExtension(request.InputPath), ".pdf", StringComparison.OrdinalIgnoreCase)) throw new ArgumentException("Redaction input must be a PDF.");
        if (request.ExternalValidators.Count > 16 || request.ExternalValidators.Any(static validator => validator is null)) throw new ArgumentException("External validator collections accept at most 16 non-null validators.");
        if (request.Recipe.SignaturePolicy != PdfRedactionSignaturePolicy.CreateAndSignDerivative && (request.OutputSigner is not null || request.OutputSignatureOptions is not null)) throw new ArgumentException("Output signer settings require CreateAndSignDerivative.");
        if (request.Recipe.Rules is null || request.Recipe.Regions is null) throw new ArgumentException("Recipe rule and region collections cannot be null.");
        if (!string.Equals(request.Recipe.Schema, PdfRedactionRecipe.CurrentSchema, StringComparison.Ordinal)) throw new ArgumentException("Unsupported redaction recipe schema.");
        if (request.Recipe.RegexTimeoutMilliseconds <= 0) throw new ArgumentOutOfRangeException(nameof(request.Recipe.RegexTimeoutMilliseconds));
        if (request.Recipe.Rules.Count > request.Limits.MaximumRules) throw new RedactionWorkflowException("Recipe rule count exceeds the configured limit.");
        if (request.Limits.MaximumInputBytes <= 0 || request.Limits.MaximumOutputBytes <= 0 || request.Limits.MaximumEvidenceBytes <= 0 || request.Limits.MaximumBatchPreparedBytes <= 0 || request.Limits.MaximumRules <= 0 || request.Limits.MaximumRuleCharacters <= 0 || request.Limits.MaximumAreas <= 0 || request.Limits.MaximumGeometryPoints <= 0 || request.Limits.MaximumCandidates <= 0 || request.Limits.MaximumBatchItems <= 0 || request.Limits.MaximumConcurrency <= 0 || request.Limits.MaximumConcurrency > 32) throw new ArgumentOutOfRangeException(nameof(request.Limits));
        long ruleCharacters = 0;
        foreach (PdfRedactionRule rule in request.Recipe.Rules) {
            if (rule is null || !Enum.IsDefined(rule.Kind)) throw new ArgumentException("Recipe rules require known kinds.");
            if (rule.Kind != PdfRedactionRuleKind.RedactAnnotations && string.IsNullOrWhiteSpace(rule.Value)) throw new ArgumentException("Literal, Regex, FormField, and LogicalKind rules require non-empty values.");
            if (rule.Kind == PdfRedactionRuleKind.RedactAnnotations && !string.IsNullOrEmpty(rule.Value)) throw new ArgumentException("RedactAnnotations does not accept a value.");
            ruleCharacters += rule.Kind.ToString().Length + (rule.Value?.Length ?? 0);
            if (ruleCharacters > request.Limits.MaximumRuleCharacters) throw new RedactionWorkflowException("Recipe rule text exceeds the configured character limit.");
        }
        ValidateRegionComplexity(request.Recipe.Regions, request.Limits);
        if (request.Mode == PdfRedactionWorkflowMode.ApplyAndVerify && string.IsNullOrWhiteSpace(request.OutputPath)) throw new ArgumentException("ApplyAndVerify requires an output path.");
        if (request.Mode == PdfRedactionWorkflowMode.VerifyExistingOutput && (string.IsNullOrWhiteSpace(request.OutputPath) || !File.Exists(request.OutputPath))) throw new FileNotFoundException("VerifyExistingOutput requires an existing output PDF.", request.OutputPath);
        ValidateArtifactPaths(request);
        if (request.Mode != PdfRedactionWorkflowMode.PlanOnly && request.Decisions is null) throw new ArgumentException("Apply and verify require a reviewed decision manifest.");
        if (request.ExpectedOutputSha256 is not null && !IsSha256(request.ExpectedOutputSha256)) throw new ArgumentException("ExpectedOutputSha256 must contain 64 hexadecimal characters.");
        if (request.Mode == PdfRedactionWorkflowMode.PlanOnly && request.Recipe.Rules.Count == 0 && request.Recipe.Regions.Count == 0) throw new ArgumentException("The recipe must contain at least one rule or region.");
        bool needsOcrTextRule = request.Recipe.DetectionMode != PdfRedactionDetectionMode.NativeOnly;
        if (needsOcrTextRule && !request.Recipe.Rules.Any(static rule => rule.Kind is PdfRedactionRuleKind.Literal or PdfRedactionRuleKind.Regex)) throw new ArgumentException("OCR detection requires at least one Literal or Regex rule.");
    }

    private static void ValidateArtifactPaths(PdfRedactionWorkflowRequest request) {
        string input = Path.GetFullPath(request.InputPath);
        string? output = NormalizeOptionalPath(request.OutputPath);
        string? evidence = NormalizeOptionalPath(request.EvidencePath);
        if (request.Mode == PdfRedactionWorkflowMode.ApplyAndVerify && output is not null && OfficeWorkflowPathIdentity.AreEquivalentWithPortableFallback(input, output)) {
            throw new ArgumentException("Redaction output must be a new artifact path and cannot replace the reviewed input in place.");
        }
        if (evidence is not null && OfficeWorkflowPathIdentity.AreEquivalentWithPortableFallback(input, evidence)) {
            throw new ArgumentException("Redaction evidence cannot replace the input PDF.");
        }
        if (evidence is not null && output is not null && OfficeWorkflowPathIdentity.AreEquivalentWithPortableFallback(output, evidence)) {
            throw new ArgumentException("Redaction output and evidence paths must be different.");
        }
        foreach (string protectedInput in request.ProtectedInputPaths) {
            if (string.IsNullOrWhiteSpace(protectedInput)) throw new ArgumentException("Protected input paths cannot be empty.");
            string protectedPath = Path.GetFullPath(protectedInput);
            if ((output is not null && OfficeWorkflowPathIdentity.AreEquivalentWithPortableFallback(output, protectedPath)) ||
                (evidence is not null && OfficeWorkflowPathIdentity.AreEquivalentWithPortableFallback(evidence, protectedPath))) {
                throw new ArgumentException("Redaction output and evidence cannot replace a protected recipe, decision, or host input path.");
            }
        }
    }

    private static void ValidateRegionComplexity(IEnumerable<PdfRedactionRecipeRegion> regions, PdfRedactionWorkflowLimits limits) {
        var stack = new Stack<(PdfRedactionRecipeRegion Region, int Depth)>();
        foreach (PdfRedactionRecipeRegion region in regions) stack.Push((region, 1));
        var visited = new HashSet<PdfRedactionRecipeRegion>(ReferenceEqualityComparer.Instance);
        long geometryEntries = 0;
        long normalizedAreas = 0;
        while (stack.Count > 0) {
            (PdfRedactionRecipeRegion region, int depth) = stack.Pop();
            if (region is null) throw new ArgumentException("Recipe regions cannot contain null entries.");
            if (region.Points is null || region.Areas is null) throw new ArgumentException("Recipe region point and area collections cannot be null.");
            if (!visited.Add(region)) throw new ArgumentException("Recipe region groups cannot contain cycles or reuse the same mutable region instance.");
            if (depth > 16) throw new RedactionWorkflowException("Recipe region nesting exceeds the supported depth of 16.");
            geometryEntries += region.Points.Count;
            geometryEntries += region.Areas.Count;
            if (!Enum.IsDefined(region.Kind)) throw new ArgumentException("Recipe regions require a known geometry kind.");
            normalizedAreas += region.Kind == PdfRedactionRegionKind.Freehand ? Math.Max(0, region.Points.Count - 1) : region.Kind == PdfRedactionRegionKind.Group ? 0 : 1;
            if (geometryEntries > limits.MaximumGeometryPoints) throw new RedactionWorkflowException("Recipe geometry exceeds the configured point limit.");
            if (normalizedAreas > limits.MaximumAreas) throw new RedactionWorkflowException("Recipe geometry exceeds the configured normalized-area limit.");
            if (region.Label?.Length > 1_024) throw new RedactionWorkflowException("Recipe region labels cannot exceed 1,024 characters.");
            foreach (PdfRedactionRecipeRegion child in region.Areas) stack.Push((child, depth + 1));
        }
    }

    private static void ValidateDecisions(PdfRedactionDecisionManifest decisions, string sourceSha, string recipeSha, IReadOnlyList<PdfRedactionWorkflowCandidate> candidates) {
        if (!string.Equals(decisions.Schema, PdfRedactionDecisionManifest.CurrentSchema, StringComparison.Ordinal)) throw new RedactionWorkflowException("Unsupported redaction decision schema.");
        if (!string.Equals(decisions.SourceSha256, sourceSha, StringComparison.Ordinal)) throw new RedactionWorkflowException("The reviewed decisions belong to different source PDF bytes.");
        if (!string.Equals(decisions.RecipeSha256, recipeSha, StringComparison.Ordinal)) throw new RedactionWorkflowException("The reviewed decisions belong to a different recipe revision.");
        if (decisions.ApprovedCandidateIds is null || decisions.RejectedCandidateIds is null) throw new RedactionWorkflowException("Decision collections cannot be null.");
        var approved = new HashSet<string>(decisions.ApprovedCandidateIds, StringComparer.Ordinal);
        var rejected = new HashSet<string>(decisions.RejectedCandidateIds, StringComparer.Ordinal);
        if (approved.Count != decisions.ApprovedCandidateIds.Count || rejected.Count != decisions.RejectedCandidateIds.Count) throw new RedactionWorkflowException("Decision manifests cannot contain duplicate candidate ids.");
        if (approved.Overlaps(rejected)) throw new RedactionWorkflowException("A redaction candidate cannot be both approved and rejected.");
        var expected = new HashSet<string>(candidates.Select(static candidate => candidate.Id), StringComparer.Ordinal);
        var decided = new HashSet<string>(approved, StringComparer.Ordinal); decided.UnionWith(rejected);
        if (!expected.SetEquals(decided)) throw new RedactionWorkflowException("The decision manifest must explicitly approve or reject every current candidate and cannot contain stale candidate ids.");
    }

    private static PdfRedactionWorkflowEvidence CreateEvidence(string sourceSha, string outputSha, string recipeSha, PdfRedactionWorkflowRequest request, CandidatePlan candidates, PdfRedactionEvidenceReport report, int ocrResidualCount, RedactionSignatureEvidence signature, IReadOnlyList<string> externalValidators) =>
        new(sourceSha, outputSha, recipeSha, request.Decisions!.ApprovedCandidateIds.Count, request.Decisions.RejectedCandidateIds.Count, report.VerifiedAbsentCount, report.ResidualCount + ocrResidualCount, report.InconclusiveCount, report.IsVerified && ocrResidualCount == 0, report.AffectedPageNumbers, report.Verification.Issues.Select(static issue => new PdfRedactionEvidenceIssue(issue.Feature)).Concat(ocrResidualCount == 0 ? Array.Empty<PdfRedactionEvidenceIssue>() : new[] { new PdfRedactionEvidenceIssue("OcrRedactionResidual") }).ToArray(), request.Recipe.EncryptedDocumentPolicy.ToString(), candidates.OcrUsed, candidates.Candidates.Where(static item => item.Provider is not null).Select(static item => item.Provider!).Distinct(StringComparer.Ordinal).ToArray(), candidates.OcrUsed, ocrResidualCount, signature.SourceCount, signature.OutputCount, request.Recipe.SignaturePolicy.ToString(), signature.SignerName, signature.CryptographicallyVerified, externalValidators);

    private static PdfRedactionWorkflowEvidence CreateVerificationEvidence(string sourceSha, string outputSha, string recipeSha, PdfRedactionWorkflowRequest request, CandidatePlan candidates, PdfRedactionPlan plan, PdfRedactionVerificationReport report, int ocrResidualCount, RedactionSignatureEvidence signature, IReadOnlyList<string> externalValidators) =>
        new(sourceSha, outputSha, recipeSha, request.Decisions!.ApprovedCandidateIds.Count, request.Decisions.RejectedCandidateIds.Count, report.IsVerified && ocrResidualCount == 0 ? plan.Matches.Count : 0, report.Issues.Count(static issue => issue.Feature == "RedactionPlanResidual") + ocrResidualCount, report.IsVerified ? 0 : plan.Matches.Count, report.IsVerified && ocrResidualCount == 0, plan.Areas.Select(static area => area.PageNumber).Distinct().OrderBy(static page => page).ToArray(), report.Issues.Select(static issue => new PdfRedactionEvidenceIssue(issue.Feature)).Concat(ocrResidualCount == 0 ? Array.Empty<PdfRedactionEvidenceIssue>() : new[] { new PdfRedactionEvidenceIssue("OcrRedactionResidual") }).ToArray(), request.Recipe.EncryptedDocumentPolicy.ToString(), candidates.OcrUsed, candidates.Candidates.Where(static item => item.Provider is not null).Select(static item => item.Provider!).Distinct(StringComparer.Ordinal).ToArray(), candidates.OcrUsed, ocrResidualCount, signature.SourceCount, signature.OutputCount, request.Recipe.SignaturePolicy.ToString(), signature.SignerName, signature.CryptographicallyVerified, externalValidators);

    private static PdfRedactionWorkflowEvidence CreateEmptyEvidence(string sourceSha, string outputSha, string recipeSha, PdfRedactionWorkflowRequest request, CandidatePlan candidates, string encryptionPolicy, RedactionSignatureEvidence signature, IReadOnlyList<string> externalValidators) =>
        new(sourceSha, outputSha, recipeSha, 0, request.Decisions!.RejectedCandidateIds.Count, 0, 0, 0, true, Array.Empty<int>(), Array.Empty<PdfRedactionEvidenceIssue>(), encryptionPolicy, candidates.OcrUsed, candidates.Candidates.Where(static item => item.Provider is not null).Select(static item => item.Provider!).Distinct(StringComparer.Ordinal).ToArray(), sourceSignatureCount: signature.SourceCount, outputSignatureCount: signature.OutputCount, signaturePolicy: request.Recipe.SignaturePolicy.ToString(), outputSigner: signature.SignerName, signatureCryptographicallyVerified: signature.CryptographicallyVerified, externalValidators: externalValidators);

    private static async Task<int> CountOcrResidualsAsync(PdfDocument output, PdfRedactionWorkflowRequest request, IReadOnlyList<PdfRedactionArea> approvedAreas, CancellationToken cancellationToken) {
        PdfRedactionSearchOptions search = BuildSearchOptions(request.Recipe, request.Limits.MaximumCandidates, cancellationToken);
        PdfOcrRedactionSearchResult post = await output.SearchRedactionCandidatesWithOcrAsync(request.OcrEngine!, search, request.OcrOptions, cancellationToken).ConfigureAwait(false);
        return post.Candidates.Count(candidate => approvedAreas.Any(area =>
            area.PageNumber == candidate.Area.PageNumber &&
            area.IntersectsRectangle(candidate.Area.X, candidate.Area.Y, candidate.Area.Width, candidate.Area.Height)));
    }

    private static PdfLoadOptions GetOutputLoadOptions(PdfRedactionWorkflowRequest request) {
        string? password = request.Recipe.EncryptedDocumentPolicy == PdfRedactionEncryptedDocumentPolicy.DecryptAndReencrypt
            ? request.OutputEncryption?.OwnerPassword ?? request.OutputEncryption?.UserPassword
            : null;
        return password is null ? PdfLoadOptions.Default : new PdfLoadOptions { Password = password };
    }

    private static void ValidateFinalArtifactEncryptionPolicy(byte[] output, PdfRedactionWorkflowRequest request, PdfLoadOptions outputOptions, CancellationToken cancellationToken) {
        PdfDocumentPreflight preflight = PdfInspector.Preflight(output, outputOptions, cancellationToken);
        bool expectedEncryption = request.Recipe.EncryptedDocumentPolicy == PdfRedactionEncryptedDocumentPolicy.DecryptAndReencrypt;
        if (preflight.Probe.HasEncryption != expectedEncryption || expectedEncryption && !preflight.Probe.Security.HasOwnerAuthorization) {
            throw new RedactionWorkflowException("The final artifact does not satisfy the selected authenticated encrypted-document policy.");
        }
        if (!expectedEncryption) return;
        PdfStandardEncryptionOptions expected = request.OutputEncryption!;
        (int version, int revision, int length) = expected.Algorithm switch {
            PdfStandardEncryptionAlgorithm.Aes256 => (5, 6, 256),
            PdfStandardEncryptionAlgorithm.Aes128 => (4, 4, 128),
            PdfStandardEncryptionAlgorithm.LegacyRc4 => (2, 3, 128),
            _ => throw new ArgumentException("Output encryption algorithm is not defined.")
        };
        PdfDocumentSecurityInfo actual = preflight.Probe.Security;
        if (!string.Equals(actual.EncryptionFilter, "Standard", StringComparison.Ordinal) ||
            actual.EncryptionVersion != version || actual.EncryptionRevision != revision || actual.EncryptionLengthBits != length ||
            actual.AllowedStandardPermissions != expected.AllowedPermissions || actual.EncryptMetadata != expected.EncryptMetadata) {
            throw new RedactionWorkflowException("The final artifact encryption algorithm, permissions, or metadata policy differs from the requested output protection contract.");
        }
    }

    private static bool IsSha256(string value) {
        if (value.Length != 64) return false;
        for (int index = 0; index < value.Length; index++) {
            char character = value[index];
            if (!((character >= '0' && character <= '9') || (character >= 'a' && character <= 'f') || (character >= 'A' && character <= 'F'))) return false;
        }
        return true;
    }

    private static PreparedRedactionResult CreatePreparedResult(PdfRedactionWorkflowResult result, string? evidencePath, PdfRedactionWorkflowLimits limits, byte[]? outputBytes, string? outputPath = null) {
        var files = new List<PreparedFile>();
        if (outputBytes is not null && outputPath is not null) files.Add(new PreparedFile(Path.GetFullPath(outputPath), outputBytes));
        if (!string.IsNullOrWhiteSpace(evidencePath)) {
            byte[] evidenceBytes = JsonSerializer.SerializeToUtf8Bytes(
                new PdfRedactionWorkflowRecord(result),
                PdfRedactionWorkflowJsonContext.Default.PdfRedactionWorkflowRecord);
            if (evidenceBytes.LongLength > limits.MaximumEvidenceBytes) throw new RedactionWorkflowException("Privacy-safe redaction evidence exceeds the configured evidence-byte limit.");
            files.Add(new PreparedFile(Path.GetFullPath(evidencePath), evidenceBytes));
        }
        return new PreparedRedactionResult(result, files);
    }

    private static PdfRedactionWorkflowResult FailedRedactionResult(PdfRedactionWorkflowRequest request, OfficeWorkflowStatus status, string summary, string code, OfficeWorkflowDiagnosticSeverity severity, Exception? exception = null) =>
        new(request.Id, request.Mode, status, summary, string.Empty, string.Empty, Array.Empty<PdfRedactionWorkflowCandidate>(), null, null, null,
            new[] { new OfficeWorkflowDiagnostic(code, SafeFailureMessage(exception, summary), severity, "redaction", exception is null ? null : new Dictionary<string, string> { ["exceptionType"] = exception.GetType().Name }) });

    private static string SafeFailureMessage(Exception? exception, string fallback) => exception switch {
        null => fallback,
        RedactionWorkflowException => exception.Message,
        OperationCanceledException => fallback,
        _ => fallback + " The privacy-safe result omits detailed exception text; inspect the exception type and host logs."
    };

    private static async Task<byte[]> ReadFileBoundedAsync(string path, long maximumBytes, CancellationToken cancellationToken) {
        var info = new FileInfo(path);
        if (info.Length > maximumBytes) throw new RedactionWorkflowException($"Input is {info.Length} bytes, above the configured {maximumBytes}-byte limit.");
        await using var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 81_920, FileOptions.Asynchronous | FileOptions.SequentialScan);
        using var output = new MemoryStream();
        byte[] buffer = new byte[81_920];
        long total = 0;
        while (true) {
            int read = await input.ReadAsync(buffer.AsMemory(0, buffer.Length), cancellationToken).ConfigureAwait(false);
            if (read == 0) break;
            total = checked(total + read);
            if (total > maximumBytes) throw new RedactionWorkflowException("Input grew above the configured limit while being read.");
            output.Write(buffer, 0, read);
        }
        return output.ToArray();
    }

    private static string ComputeRecipeSha256(PdfRedactionRecipe recipe) => ComputeSha256(JsonSerializer.SerializeToUtf8Bytes(
        recipe,
        PdfRedactionWorkflowJsonContext.Default.PdfRedactionRecipe));
    private static string ComputeSha256(byte[] bytes) => System.Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    private static string ComputeCandidateId(string sourceSha, string origin, IReadOnlyList<PdfRedactionArea> areas) {
        var identity = new StringBuilder(sourceSha).Append('|').Append(origin);
        foreach (PdfRedactionArea area in areas) {
            identity.Append('|').Append(area.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture))
                .Append('|').Append(area.X.ToString("R", System.Globalization.CultureInfo.InvariantCulture))
                .Append('|').Append(area.Y.ToString("R", System.Globalization.CultureInfo.InvariantCulture))
                .Append('|').Append(area.Width.ToString("R", System.Globalization.CultureInfo.InvariantCulture))
                .Append('|').Append(area.Height.ToString("R", System.Globalization.CultureInfo.InvariantCulture))
                .Append('|').Append(area.Label ?? string.Empty);
            if (area.ExactGeometry is PdfRedactionGeometry geometry) {
                identity.Append('|').Append(geometry.Kind).Append('|')
                    .Append(geometry.StrokeWidth.ToString("R", System.Globalization.CultureInfo.InvariantCulture));
                foreach (PdfRedactionPoint point in geometry.Points) {
                    identity.Append('|').Append(point.X.ToString("R", System.Globalization.CultureInfo.InvariantCulture))
                        .Append(',').Append(point.Y.ToString("R", System.Globalization.CultureInfo.InvariantCulture));
                }
            }
        }
        return ComputeSha256(Encoding.UTF8.GetBytes(identity.ToString()));
    }
    private static string? NormalizeOptionalPath(string? path) => string.IsNullOrWhiteSpace(path) ? null : Path.GetFullPath(path);

    private static long GetMaximumPreparedBytes(PdfRedactionWorkflowRequest request) {
        long maximum = 0;
        if (request.Mode == PdfRedactionWorkflowMode.ApplyAndVerify && request.OutputPath is not null) maximum = checked(maximum + request.Limits.MaximumOutputBytes);
        if (request.EvidencePath is not null) maximum = checked(maximum + request.Limits.MaximumEvidenceBytes);
        if (maximum > request.Limits.MaximumBatchPreparedBytes) {
            throw new RedactionWorkflowException("An item\'s configured output and evidence ceilings exceed the atomic batch prepared-byte budget.");
        }
        return maximum;
    }

    private static void DisposeReservations(IEnumerable<PreparedByteReservation?> reservations) {
        foreach (PreparedByteReservation? reservation in reservations) reservation?.Dispose();
    }

    private static PdfRedactionWorkflowResult MarkNotPublished(PdfRedactionWorkflowResult result, OfficeWorkflowStatus status, string summary, Exception? exception = null) {
        var diagnostics = result.Diagnostics.Concat(new[] {
            new OfficeWorkflowDiagnostic("RedactionBatchNotPublished", SafeFailureMessage(exception, summary),
                status == OfficeWorkflowStatus.Cancelled ? OfficeWorkflowDiagnosticSeverity.Information : OfficeWorkflowDiagnosticSeverity.Error,
                "redaction",
                exception is null ? null : new Dictionary<string, string> { ["exceptionType"] = exception.GetType().Name })
        }).ToArray();
        return new PdfRedactionWorkflowResult(result.RequestId, result.Mode, status, summary, result.SourceSha256, result.RecipeSha256, result.Candidates, null, null, null, diagnostics);
    }

    private static PdfRedactionWorkflowResult MarkPublishedWithCleanupFailure(PdfRedactionWorkflowResult result, Exception exception) {
        const string summary = "The artifact was published, but prior-destination rollback data could not be removed. Host cleanup is required.";
        var diagnostics = result.Diagnostics.Concat(new[] {
            new OfficeWorkflowDiagnostic("RedactionRollbackCleanupFailed", SafeFailureMessage(exception, summary), OfficeWorkflowDiagnosticSeverity.Error, "redaction",
                new Dictionary<string, string> { ["exceptionType"] = exception.GetType().Name })
        }).ToArray();
        return new PdfRedactionWorkflowResult(result.RequestId, result.Mode, OfficeWorkflowStatus.Failed, summary, result.SourceSha256, result.RecipeSha256, result.Candidates, result.OutputPath, result.EvidencePath, result.Evidence, diagnostics);
    }

    private static void EnsureUniqueDestinations(IReadOnlyList<PreparedFile> files) {
        for (int left = 0; left < files.Count; left++) {
            for (int right = left + 1; right < files.Count; right++) {
                if (OfficeWorkflowPathIdentity.AreEquivalentWithPortableFallback(files[left].Path, files[right].Path)) {
                    throw new RedactionWorkflowException("Atomic batch output and evidence destinations must be unique.");
                }
            }
        }
    }

    private static void EnsureDestinationsDoNotReplaceReviewedInputs(IReadOnlyList<PreparedFile> files, IReadOnlyList<PdfRedactionWorkflowRequest> requests) {
        foreach (PreparedFile file in files) {
            foreach (PdfRedactionWorkflowRequest request in requests) {
                bool replacesSource = OfficeWorkflowPathIdentity.AreEquivalentWithPortableFallback(file.Path, request.InputPath);
                bool replacesVerifiedOutput = request.Mode == PdfRedactionWorkflowMode.VerifyExistingOutput &&
                    request.OutputPath is not null &&
                    OfficeWorkflowPathIdentity.AreEquivalentWithPortableFallback(file.Path, request.OutputPath);
                bool replacesProtectedInput = request.ProtectedInputPaths.Any(path => OfficeWorkflowPathIdentity.AreEquivalentWithPortableFallback(file.Path, path));
                if (replacesSource || replacesVerifiedOutput || replacesProtectedInput) {
                    throw new RedactionWorkflowException("Atomic batch destinations cannot replace any source or existing output inspected by the batch.");
                }
            }
        }
    }

    private static void PublishPreparedFiles(IReadOnlyList<PreparedFile> files, OfficeWorkflowConflictPolicy conflictPolicy, CancellationToken cancellationToken) {
        EnsureUniqueDestinations(files);
        var staged = new List<(string Temp, string Destination)>(files.Count);
        var published = new List<(string Destination, string? Backup)>(files.Count);
        bool committed = false;
        try {
            foreach (PreparedFile file in files) {
                cancellationToken.ThrowIfCancellationRequested();
                string? directory = Path.GetDirectoryName(file.Path);
                if (string.IsNullOrWhiteSpace(directory)) throw new ArgumentException("Output path requires a directory.");
                Directory.CreateDirectory(directory);
                if (conflictPolicy == OfficeWorkflowConflictPolicy.Fail && File.Exists(file.Path)) throw new IOException("Output already exists: " + file.Path);
                if (conflictPolicy == OfficeWorkflowConflictPolicy.Rename && File.Exists(file.Path)) throw new NotSupportedException("Redaction transaction publication requires Fail or Replace conflict policy.");
                string temp = Path.Combine(directory, "." + Path.GetFileName(file.Path) + "." + Guid.NewGuid().ToString("N") + ".tmp");
                staged.Add((temp, file.Path));
                File.WriteAllBytes(temp, file.Bytes);
            }
            foreach ((string temp, string destination) in staged) {
                cancellationToken.ThrowIfCancellationRequested();
                if (conflictPolicy == OfficeWorkflowConflictPolicy.Fail) {
                    File.Move(temp, destination);
                    published.Add((destination, null));
                    continue;
                }

                if (File.Exists(destination)) {
                    string backup = destination + "." + Guid.NewGuid().ToString("N") + ".rollback";
                    File.Move(destination, backup);
                    published.Add((destination, backup));
                    File.Move(temp, destination);
                } else {
                    File.Move(temp, destination);
                    published.Add((destination, null));
                }
            }
            committed = true;
        } catch {
            var rollbackFailures = new List<Exception>();
            for (int index = published.Count - 1; index >= 0; index--) {
                (string destination, string? backup) = published[index];
                try {
                    if (File.Exists(destination)) File.Delete(destination);
                    if (backup is not null && File.Exists(backup)) File.Move(backup, destination);
                } catch (Exception rollbackException) when (rollbackException is IOException or UnauthorizedAccessException) {
                    rollbackFailures.Add(rollbackException);
                }
            }
            if (rollbackFailures.Count > 0) throw new AggregateException("Redaction publication failed and one or more destinations require recovery from their .rollback sibling files.", rollbackFailures);
            throw;
        } finally {
            foreach ((string temp, _) in staged) if (File.Exists(temp)) File.Delete(temp);
        }
        if (committed) {
            var cleanupFailures = new List<Exception>();
            foreach ((_, string? backup) in published) {
                if (backup is null || !File.Exists(backup)) continue;
                try { File.Delete(backup); } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { cleanupFailures.Add(exception); }
            }
            if (cleanupFailures.Count > 0) throw new RedactionBackupCleanupException(cleanupFailures);
        }
    }

    private sealed record CandidateOrigin(string Origin, double? Confidence, string? Provider, string? Model, string? Language);
    private sealed class RedactionWorkflowException : InvalidOperationException { internal RedactionWorkflowException(string message) : base(message) { } }
    private sealed class RedactionBackupCleanupException : IOException {
        internal RedactionBackupCleanupException(IReadOnlyCollection<Exception> failures)
            : base($"{failures.Count} rollback file(s) could not be removed after publication.", failures.FirstOrDefault()) { }
    }

    private sealed record CandidatePlan(IReadOnlyList<IReadOnlyList<PdfRedactionArea>> CandidateAreas, IReadOnlyList<PdfRedactionWorkflowCandidate> Candidates, IReadOnlyList<OfficeWorkflowDiagnostic> Diagnostics, bool OcrUsed) {
        internal int AreaCount => CandidateAreas.Sum(static areas => areas.Count);
    }
    private sealed record PreparedFile(string Path, byte[] Bytes);
    private sealed record PreparedRedactionResult(PdfRedactionWorkflowResult Result, IReadOnlyList<PreparedFile> Files);
}
