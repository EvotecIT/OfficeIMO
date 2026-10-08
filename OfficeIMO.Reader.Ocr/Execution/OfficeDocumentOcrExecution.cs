using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Ocr;

namespace OfficeIMO.Reader;

/// <summary>Runs optional OCR engines against bounded Reader candidates and merges recognized text.</summary>
public static partial class OfficeDocumentOcrExecutionExtensions {
    /// <summary>
    /// Executes an OCR engine over validated candidate assets, preserves deterministic result order, and enriches the document.
    /// </summary>
    public static async Task<OfficeDocumentOcrExecutionResult> ApplyOcrAsync(
        this OfficeDocumentReadResult document,
        IOcrEngine engine,
        OfficeDocumentOcrExecutionOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (engine == null) throw new ArgumentNullException(nameof(engine));

        ExecutionOptionsSnapshot effective = ExecutionOptionsSnapshot.Create(options);
        return await ApplyOcrDocumentAsync(document, OcrEngineRunner.CreateExecution(engine), effective,
            new ExecutionBudget(effective), new TimedOutOcrOperationTracker(), "root", document.Source?.Path,
            cancellationToken).ConfigureAwait(false);
    }

    private static async Task<OfficeDocumentOcrExecutionResult> ApplyOcrDocumentAsync(
        OfficeDocumentReadResult document, OcrEngineExecution engineExecution, ExecutionOptionsSnapshot effective,
        ExecutionBudget budget, TimedOutOcrOperationTracker timedOutOperations, string documentId,
        string? documentPath, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        string engineId = engineExecution.Id;
        IReadOnlyList<OfficeDocumentOcrCandidate> candidates = OfficeDocumentOcrCandidates.Collect(document);
        IReadOnlyList<OfficeDocumentAsset> assets = (document.Assets ?? Array.Empty<OfficeDocumentAsset>())
            .Concat((document.Pages ?? Array.Empty<OfficeDocumentPage>()).SelectMany(page => page.Assets ?? Array.Empty<OfficeDocumentAsset>())).ToArray();
        OcrEngineCapabilities capabilities = engineExecution.Capabilities;
        var diagnostics = new List<OfficeDocumentDiagnostic>();
        int selectedCount = Math.Min(candidates.Count, effective.MaxCandidates - budget.SelectedCandidates);
        List<CandidateJob> jobs = BuildJobs(document, candidates, assets, capabilities, engineId, effective, budget, diagnostics, cancellationToken);
        int degree = capabilities.SupportsConcurrentRequests ? Math.Min(effective.MaxDegreeOfParallelism, Math.Max(1, jobs.Count)) : 1;

        var recognitions = new List<OfficeDocumentOcrRecognition>(jobs.Count);
        var recognizedText = new List<OfficeDocumentOcrTextResult>(jobs.Count);
        int failedCount = 0;
        int emptyCount = 0;
        int attemptedCount = 0;
        long inputBytes = 0;
        void Accept(CandidateOutcome outcome) {
            cancellationToken.ThrowIfCancellationRequested();
            if (outcome.WasAttempted) inputBytes += outcome.Job.Payload.LongLength;
            if (outcome.WasAttempted) attemptedCount++;
            if (outcome.FailureDiagnostic != null) {
                if (outcome.WasAttempted) failedCount++;
                diagnostics.Add(outcome.FailureDiagnostic);
                return;
            }

            OcrResult engineResult = outcome.Result ?? new OcrResult();
            int diagnosticStart = diagnostics.Count;
            NormalizeEngineResult(engineResult, engineId, effective, budget, outcome.Job.Candidate, diagnostics, cancellationToken);
            bool normalizationLoss = diagnostics.Skip(diagnosticStart).Any(item => item.Severity != OfficeDocumentDiagnosticSeverity.Information);
            budget.Consume(engineResult);
            recognitions.Add(new OfficeDocumentOcrRecognition {
                DocumentId = documentId,
                DocumentPath = documentPath,
                CandidateId = outcome.Job.Candidate.Id,
                AssetId = outcome.Job.Asset.Id,
                Result = engineResult
            });
            diagnostics.AddRange((engineResult.Diagnostics ?? Array.Empty<OcrDiagnostic>())
                .Where(static diagnostic => diagnostic != null)
                .Select(diagnostic => MapProviderDiagnostic(diagnostic, engineId, outcome.Job.Candidate)));
            if (string.IsNullOrWhiteSpace(engineResult.Text)) {
                emptyCount++;
                diagnostics.Add(BuildDiagnostic(
                    outcome.Job.Candidate,
                    outcome.Job.Asset,
                    engineId,
                    OfficeDocumentDiagnosticSeverity.Warning,
                    OfficeDocumentDiagnosticCategory.Ocr,
                    "ocr-empty-result",
                    "The OCR engine completed without recognized text.",
                    true));
                return;
            }

            recognizedText.Add(new OfficeDocumentOcrTextResult {
                CandidateId = outcome.Job.Candidate.Id,
                Text = engineResult.Text,
                Confidence = engineResult.Confidence,
                Language = engineResult.Language,
                Provider = engineResult.Provider,
                Model = engineResult.Model,
                Recognition = CaptureRecognitionEvidence(engineResult, normalizationLoss)
            });
        }

        await ExecuteCandidatesAsync(document, engineExecution, engineId, jobs, effective, budget, degree,
            timedOutOperations, documentPath, Accept, cancellationToken).ConfigureAwait(false);

        OfficeDocumentOcrEnrichmentResult enrichment = document.ApplyOcrResults(recognizedText, effective.EnrichmentOptions);
        OfficeDocumentReadResult enriched = enrichment.Document;
        enriched.CapabilitiesUsed = AppendCapabilities(enriched.CapabilitiesUsed, engineId);
        enriched.Diagnostics = (enriched.Diagnostics ?? Array.Empty<OfficeDocumentDiagnostic>()).Concat(diagnostics).ToArray();
        int skippedCount = candidates.Count - attemptedCount;
        var report = new OfficeDocumentOcrExecutionReport {
            EngineId = engineId,
            DocumentCount = 1,
            CandidateCount = candidates.Count,
            SelectedCandidateCount = selectedCount,
            AttemptedCandidateCount = attemptedCount,
            RecognizedCandidateCount = recognizedText.Count,
            EmptyCandidateCount = emptyCount,
            SkippedCandidateCount = skippedCount,
            FailedCandidateCount = failedCount,
            LineSpanCount = CountSpans(recognitions, OcrTextSpanLevel.Line),
            WordSpanCount = CountSpans(recognitions, OcrTextSpanLevel.Word),
            CharacterSpanCount = CountSpans(recognitions, OcrTextSpanLevel.Character),
            InputBytes = inputBytes,
            EffectiveDegreeOfParallelism = attemptedCount == 0 ? 0 : degree
        };
        enriched.Metadata = BuildExecutionMetadata(enriched.Metadata, report);

        return new OfficeDocumentOcrExecutionResult {
            Document = enriched,
            Recognitions = recognitions,
            Diagnostics = diagnostics,
            Report = report
        };
    }

    private static async Task ExecuteCandidatesAsync(
        OfficeDocumentReadResult document, OcrEngineExecution engineExecution, string engineId,
        IReadOnlyList<CandidateJob> jobs, ExecutionOptionsSnapshot options, ExecutionBudget budget,
        int degree, TimedOutOcrOperationTracker timedOutOperations, string? documentPath, Action<CandidateOutcome> accept,
        CancellationToken cancellationToken) {
        var running = new List<Task<CandidateOutcome>>(degree);
        int nextJob = 0;
        try {
            while (nextJob < jobs.Count || running.Count > 0) {
                cancellationToken.ThrowIfCancellationRequested();
                Task<CandidateOutcome>? fault = running.FirstOrDefault(task => task.IsFaulted || task.IsCanceled);
                if (fault != null) await fault.ConfigureAwait(false);
                while (nextJob < jobs.Count && running.Count < degree) {
                    if (!budget.CanRecognize) {
                        CandidateJob skipped = jobs[nextJob++];
                        accept(CandidateOutcome.Skipped(skipped, BuildDiagnostic(skipped.Candidate, skipped.Asset, engineId,
                            OfficeDocumentDiagnosticSeverity.Warning, OfficeDocumentDiagnosticCategory.Limit,
                            budget.RemainingTime <= TimeSpan.Zero ? "ocr-total-time-limit" : "ocr-total-text-limit",
                            "OCR was not started because the execution's total time or recognized-text budget was reached.", true)));
                    } else {
                        running.Add(ExecuteCandidateAsync(document, engineExecution, engineId, jobs[nextJob++], options,
                            budget, timedOutOperations, documentPath, cancellationToken));
                    }
                }
                if (running.Count == 0) continue;
                // Bound retained raw responses to the in-flight window; normalize in source order.
                CandidateOutcome outcome = await running[0].ConfigureAwait(false);
                running.RemoveAt(0);
                accept(outcome);
            }
        } catch {
            await AwaitRemainingCandidatesAsync(running).ConfigureAwait(false);
            throw;
        }
    }

    private static async Task AwaitRemainingCandidatesAsync(IEnumerable<Task<CandidateOutcome>> running) {
        try {
            await Task.WhenAll(running).ConfigureAwait(false);
        } catch {
            // Preserve the first fail-fast exception after observing every already-started candidate.
        }
    }

    private static async Task<CandidateOutcome> ExecuteCandidateAsync(
        OfficeDocumentReadResult document,
        OcrEngineExecution engineExecution,
        string engineId,
        CandidateJob job,
        ExecutionOptionsSnapshot options,
        ExecutionBudget budget,
        TimedOutOcrOperationTracker timedOutOperations,
        string? documentPath, CancellationToken cancellationToken) {
        try {
            cancellationToken.ThrowIfCancellationRequested();
            if (timedOutOperations.HasTimedOutOperation) {
                return CandidateOutcome.Skipped(job, BuildDiagnostic(
                    job.Candidate,
                    job.Asset,
                    engineId,
                    OfficeDocumentDiagnosticSeverity.Error,
                    OfficeDocumentDiagnosticCategory.Ocr,
                    "ocr-engine-timeout",
                    "OCR was not started because an earlier provider call exceeded its execution timeout.",
                    true));
            }
            var request = new OcrRequest {
                Payload = job.Payload,
                MediaType = job.Asset.MediaType ?? string.Empty,
                FileName = job.Asset.FileName,
                SourceId = document.Source?.SourceId,
                SourceName = documentPath,
                CandidateId = job.Candidate.Id,
                CandidateKind = job.Candidate.Kind,
                PageNumber = job.Candidate.Location?.Page,
                PixelWidth = job.Asset.Width,
                PixelHeight = job.Asset.Height,
                Region = ToOcrRegion(job.Candidate.Region),
                RegionCoordinateUnit = OcrCoordinateUnit.Points,
                Language = options.Language,
                ProviderOptions = options.ProviderOptions
            };
            TimeSpan timeout = budget.RemainingTime;
            bool totalDeadline = timeout <= options.CandidateTimeout;
            if (timeout > options.CandidateTimeout) timeout = options.CandidateTimeout;
            if (timeout <= TimeSpan.Zero) timeout = TimeSpan.FromTicks(1);
            try {
                // Each retained diagnostic can have one empty key; other unique keys consume characters.
                OcrResult result = await engineExecution.RecognizeAsync(
                    request,
                    timeout,
                    new OcrResultCaptureLimits(options.MaxSpansPerCandidate, options.MaxProviderDiagnosticsPerCandidate,
                        (int)Math.Min(options.MaxProviderDiagnosticAttributesPerCandidate, (long)options.MaxProviderDiagnosticAttributeCharactersPerCandidate + options.MaxProviderDiagnosticsPerCandidate)),
                    cancellationToken).ConfigureAwait(false);
                return CandidateOutcome.Success(job, result);
            } catch (OcrEngineTimeoutException exception) {
                timedOutOperations.MarkTimedOut();
                if (totalDeadline) budget.MarkDeadlineExpired();
                if (options.ContinueOnError) {
                    OfficeDocumentDiagnostic diagnostic = BuildDiagnostic(
                        job.Candidate,
                        job.Asset,
                        engineId,
                        OfficeDocumentDiagnosticSeverity.Error,
                        OfficeDocumentDiagnosticCategory.Ocr,
                        totalDeadline ? "ocr-total-time-limit" : "ocr-engine-timeout",
                        totalDeadline
                            ? "OCR exceeded the execution TotalTimeout (" + options.TotalTimeout + ")."
                            : exception.ProviderCallStarted
                            ? "OCR engine exceeded CandidateTimeout (" + options.CandidateTimeout + ")."
                            : "OCR was not started because the shared engine remained busy through CandidateTimeout (" + options.CandidateTimeout + ").",
                        true);
                    return exception.ProviderCallStarted
                        ? CandidateOutcome.Failure(job, diagnostic)
                        : CandidateOutcome.Skipped(job, diagnostic);
                }
                throw;
            }
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            throw;
        } catch (Exception exception) when (options.ContinueOnError && exception is not OutOfMemoryException && exception is not StackOverflowException) {
            return CandidateOutcome.Failure(job, BuildDiagnostic(
                job.Candidate,
                job.Asset,
                engineId,
                OfficeDocumentDiagnosticSeverity.Error,
                OfficeDocumentDiagnosticCategory.Ocr,
                "ocr-engine-failed",
                "OCR engine failed for this candidate. Check the configured provider and its private logs.",
                true,
                new Dictionary<string, string>(StringComparer.Ordinal) { ["exceptionType"] = exception.GetType().FullName ?? exception.GetType().Name }));
        }
    }

    private static IReadOnlyList<string> AppendCapabilities(IReadOnlyList<string>? existing, string engineId) {
        return (existing ?? Array.Empty<string>())
            .Concat(new[] { "officeimo.reader.ocr-execution", "officeimo.reader.ocr-engine." + NormalizeCapabilityToken(engineId) })
            .Where(static value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.Ordinal)
            .ToArray();
    }

    private static string NormalizeCapabilityToken(string value) {
        char[] chars = value.Trim().ToLowerInvariant().Select(static ch => char.IsLetterOrDigit(ch) ? ch : '-').ToArray();
        return new string(chars).Trim('-');
    }

    private static IReadOnlyList<OfficeDocumentMetadataEntry> BuildExecutionMetadata(
        IReadOnlyList<OfficeDocumentMetadataEntry>? existing,
        OfficeDocumentOcrExecutionReport report) {
        var entries = (existing ?? Array.Empty<OfficeDocumentMetadataEntry>())
            .Where(static entry => !entry.Id.StartsWith("reader-ocr-execution-", StringComparison.Ordinal))
            .ToList();
        AddExecutionCount(entries, "document-count", "DocumentCount", report.DocumentCount);
        AddExecutionCount(entries, "skipped-document-count", "SkippedDocumentCount", report.SkippedDocumentCount);
        AddExecutionCount(entries, "candidate-count", "CandidateCount", report.CandidateCount);
        AddExecutionCount(entries, "attempted-count", "AttemptedCount", report.AttemptedCandidateCount);
        AddExecutionCount(entries, "recognized-count", "RecognizedCount", report.RecognizedCandidateCount);
        AddExecutionCount(entries, "empty-count", "EmptyCount", report.EmptyCandidateCount);
        AddExecutionCount(entries, "skipped-count", "SkippedCount", report.SkippedCandidateCount);
        AddExecutionCount(entries, "failed-count", "FailedCount", report.FailedCandidateCount);
        AddExecutionCount(entries, "line-span-count", "LineSpanCount", report.LineSpanCount);
        AddExecutionCount(entries, "word-span-count", "WordSpanCount", report.WordSpanCount);
        AddExecutionCount(entries, "character-span-count", "CharacterSpanCount", report.CharacterSpanCount);
        entries.Add(new OfficeDocumentMetadataEntry {
            Id = "reader-ocr-execution-input-bytes",
            Category = "reader.ocr.execution",
            Name = "InputBytes",
            Value = report.InputBytes.ToString(CultureInfo.InvariantCulture),
            ValueType = "number",
            Attributes = new Dictionary<string, string>(StringComparer.Ordinal) { ["engine"] = report.EngineId }
        });
        return entries;
    }

    private static void AddExecutionCount(List<OfficeDocumentMetadataEntry> entries, string idSuffix, string name, int value) {
        entries.Add(new OfficeDocumentMetadataEntry {
            Id = "reader-ocr-execution-" + idSuffix,
            Category = "reader.ocr.execution",
            Name = name,
            Value = value.ToString(CultureInfo.InvariantCulture),
            ValueType = "count"
        });
    }

    private static int CountSpans(IReadOnlyList<OfficeDocumentOcrRecognition> recognitions, OcrTextSpanLevel level) {
        return recognitions.Sum(recognition => (recognition.Result.Spans ?? Array.Empty<OcrTextSpan>()).Count(span => span.Level == level));
    }

    private sealed class CandidateJob {
        internal CandidateJob(int index, OfficeDocumentOcrCandidate candidate, OfficeDocumentAsset asset, byte[] payload) {
            Index = index;
            Candidate = candidate;
            Asset = asset;
            Payload = payload;
        }

        internal int Index { get; }
        internal OfficeDocumentOcrCandidate Candidate { get; }
        internal OfficeDocumentAsset Asset { get; }
        internal byte[] Payload { get; }
    }

    private sealed class CandidateOutcome {
        private CandidateOutcome(CandidateJob job, OcrResult? result, OfficeDocumentDiagnostic? failureDiagnostic, bool wasAttempted) {
            Job = job;
            Result = result;
            FailureDiagnostic = failureDiagnostic;
            WasAttempted = wasAttempted;
        }

        internal CandidateJob Job { get; }
        internal OcrResult? Result { get; }
        internal OfficeDocumentDiagnostic? FailureDiagnostic { get; }
        internal bool WasAttempted { get; }

        internal static CandidateOutcome Success(CandidateJob job, OcrResult result) => new CandidateOutcome(job, result, null, true);
        internal static CandidateOutcome Failure(CandidateJob job, OfficeDocumentDiagnostic diagnostic) => new CandidateOutcome(job, null, diagnostic, true);
        internal static CandidateOutcome Skipped(CandidateJob job, OfficeDocumentDiagnostic diagnostic) => new CandidateOutcome(job, null, diagnostic, false);
    }

}
