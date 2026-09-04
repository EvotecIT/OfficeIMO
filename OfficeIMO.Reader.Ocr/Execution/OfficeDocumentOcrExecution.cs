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
        OcrEngineExecution engineExecution = OcrEngineRunner.CreateExecution(engine);
        string engineId = engineExecution.Id;
        IReadOnlyList<OfficeDocumentOcrCandidate> candidates = document.OcrCandidates ?? Array.Empty<OfficeDocumentOcrCandidate>();
        IReadOnlyList<OfficeDocumentAsset> assets = document.Assets ?? Array.Empty<OfficeDocumentAsset>();
        OcrEngineCapabilities capabilities = engineExecution.Capabilities;
        var diagnostics = new List<OfficeDocumentDiagnostic>();
        List<CandidateJob> jobs = BuildJobs(document, candidates, assets, capabilities, engineId, effective, diagnostics);
        int degree = capabilities.SupportsConcurrentRequests ? Math.Min(effective.MaxDegreeOfParallelism, Math.Max(1, jobs.Count)) : 1;

        var timedOutOperations = new TimedOutOcrOperationTracker();
        CandidateOutcome[] outcomes = await ExecuteCandidatesAsync(
            document,
            engineExecution,
            engineId,
            jobs,
            effective,
            degree,
            timedOutOperations,
            cancellationToken).ConfigureAwait(false);

        var recognitions = new List<OfficeDocumentOcrRecognition>(outcomes.Length);
        var recognizedText = new List<OfficeDocumentOcrTextResult>(outcomes.Length);
        int failedCount = 0;
        int emptyCount = 0;
        int attemptedCount = 0;
        foreach (CandidateOutcome outcome in outcomes.OrderBy(static outcome => outcome.Job.Index)) {
            if (outcome.WasAttempted) attemptedCount++;
            if (outcome.FailureDiagnostic != null) {
                if (outcome.WasAttempted) failedCount++;
                diagnostics.Add(outcome.FailureDiagnostic);
                continue;
            }

            OcrResult engineResult = outcome.Result ?? new OcrResult();
            NormalizeEngineResult(engineResult, engineId, effective, outcome.Job.Candidate, diagnostics);
            recognitions.Add(new OfficeDocumentOcrRecognition {
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
                continue;
            }

            recognizedText.Add(new OfficeDocumentOcrTextResult {
                CandidateId = outcome.Job.Candidate.Id,
                Text = engineResult.Text,
                Confidence = engineResult.Confidence,
                Language = engineResult.Language,
                Provider = engineResult.Provider,
                Model = engineResult.Model
            });
        }

        OfficeDocumentOcrEnrichmentResult enrichment = document.ApplyOcrResults(recognizedText, effective.EnrichmentOptions);
        OfficeDocumentReadResult enriched = enrichment.Document;
        enriched.CapabilitiesUsed = AppendCapabilities(enriched.CapabilitiesUsed, engineId);
        enriched.Diagnostics = (enriched.Diagnostics ?? Array.Empty<OfficeDocumentDiagnostic>()).Concat(diagnostics).ToArray();
        int skippedCount = candidates.Count - attemptedCount;
        var report = new OfficeDocumentOcrExecutionReport {
            EngineId = engineId,
            CandidateCount = candidates.Count,
            SelectedCandidateCount = Math.Min(candidates.Count, effective.MaxCandidates),
            AttemptedCandidateCount = attemptedCount,
            RecognizedCandidateCount = recognizedText.Count,
            EmptyCandidateCount = emptyCount,
            SkippedCandidateCount = skippedCount,
            FailedCandidateCount = failedCount,
            LineSpanCount = CountSpans(recognitions, OcrTextSpanLevel.Line),
            WordSpanCount = CountSpans(recognitions, OcrTextSpanLevel.Word),
            CharacterSpanCount = CountSpans(recognitions, OcrTextSpanLevel.Character),
            InputBytes = outcomes.Where(static outcome => outcome.WasAttempted).Sum(static outcome => (long)outcome.Job.Payload.LongLength),
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

    private static async Task<CandidateOutcome[]> ExecuteCandidatesAsync(
        OfficeDocumentReadResult document,
        OcrEngineExecution engineExecution,
        string engineId,
        IReadOnlyList<CandidateJob> jobs,
        ExecutionOptionsSnapshot options,
        int degree,
        TimedOutOcrOperationTracker timedOutOperations,
        CancellationToken cancellationToken) {
        var outcomes = new List<CandidateOutcome>(jobs.Count);
        var running = new List<Task<CandidateOutcome>>(degree);
        int nextJob = 0;

        while (nextJob < jobs.Count || running.Count > 0) {
            while (nextJob < jobs.Count && running.Count < degree) {
                running.Add(ExecuteCandidateAsync(document, engineExecution, engineId, jobs[nextJob++], options,
                    timedOutOperations, cancellationToken));
            }

            await Task.WhenAny(running).ConfigureAwait(false);
            Task<CandidateOutcome>[] completed = running
                .Where(static task => task.IsCompleted)
                .OrderBy(static task => task.IsFaulted ? 0 : 1)
                .ToArray();
            try {
                foreach (Task<CandidateOutcome> task in completed) {
                    running.Remove(task);
                    outcomes.Add(await task.ConfigureAwait(false));
                }
            } catch {
                await AwaitRemainingCandidatesAsync(running).ConfigureAwait(false);
                throw;
            }
        }

        return outcomes.ToArray();
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
        TimedOutOcrOperationTracker timedOutOperations,
        CancellationToken cancellationToken) {
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
                SourceName = document.Source?.Path,
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
            try {
                OcrResult result = await engineExecution.RecognizeAsync(
                    request,
                    options.CandidateTimeout,
                    cancellationToken).ConfigureAwait(false);
                return CandidateOutcome.Success(job, result);
            } catch (OcrEngineTimeoutException exception) {
                timedOutOperations.MarkTimedOut();
                if (options.ContinueOnError) {
                    OfficeDocumentDiagnostic diagnostic = BuildDiagnostic(
                        job.Candidate,
                        job.Asset,
                        engineId,
                        OfficeDocumentDiagnosticSeverity.Error,
                        OfficeDocumentDiagnosticCategory.Ocr,
                        "ocr-engine-timeout",
                        exception.ProviderCallStarted
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
        } catch (Exception exception) when (options.ContinueOnError) {
            return CandidateOutcome.Failure(job, BuildDiagnostic(
                job.Candidate,
                job.Asset,
                engineId,
                OfficeDocumentDiagnosticSeverity.Error,
                OfficeDocumentDiagnosticCategory.Ocr,
                "ocr-engine-failed",
                "OCR engine failed for candidate '" + job.Candidate.Id + "': " + exception.Message,
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

    private sealed class ExecutionOptionsSnapshot {
        private ExecutionOptionsSnapshot() { }

        internal string? Language { get; private set; }
        internal int MaxCandidates { get; private set; }
        internal long MaxInputBytesPerCandidate { get; private set; }
        internal long MaxTotalInputBytes { get; private set; }
        internal int MaxDegreeOfParallelism { get; private set; }
        internal TimeSpan CandidateTimeout { get; private set; }
        internal int MaxRecognizedCharactersPerCandidate { get; private set; }
        internal int MaxSpansPerCandidate { get; private set; }
        internal int MaxSpanCharactersPerCandidate { get; private set; }
        internal int MaxResultMetadataCharactersPerCandidate { get; private set; }
        internal int MaxProviderDiagnosticsPerCandidate { get; private set; }
        internal int MaxProviderDiagnosticCharactersPerCandidate { get; private set; }
        internal int MaxProviderDiagnosticAttributesPerCandidate { get; private set; }
        internal int MaxProviderDiagnosticAttributeCharactersPerCandidate { get; private set; }
        internal bool ContinueOnError { get; private set; }
        internal bool RequirePayloadHashMatch { get; private set; }
        internal IReadOnlyDictionary<string, string> ProviderOptions { get; private set; } = new Dictionary<string, string>(StringComparer.Ordinal);
        internal OfficeDocumentOcrEnrichmentOptions EnrichmentOptions { get; private set; } = new OfficeDocumentOcrEnrichmentOptions();

        internal static ExecutionOptionsSnapshot Create(OfficeDocumentOcrExecutionOptions? options) {
            OfficeDocumentOcrExecutionOptions source = options ?? new OfficeDocumentOcrExecutionOptions();
            if (source.MaxCandidates < 1) throw new ArgumentOutOfRangeException(nameof(source.MaxCandidates));
            if (source.MaxInputBytesPerCandidate < 1) throw new ArgumentOutOfRangeException(nameof(source.MaxInputBytesPerCandidate));
            if (source.MaxTotalInputBytes < 1) throw new ArgumentOutOfRangeException(nameof(source.MaxTotalInputBytes));
            if (source.MaxDegreeOfParallelism < 1) throw new ArgumentOutOfRangeException(nameof(source.MaxDegreeOfParallelism));
            if (source.CandidateTimeout <= TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(source.CandidateTimeout));
            if (source.MaxRecognizedCharactersPerCandidate < 1) throw new ArgumentOutOfRangeException(nameof(source.MaxRecognizedCharactersPerCandidate));
            if (source.MaxSpansPerCandidate < 1) throw new ArgumentOutOfRangeException(nameof(source.MaxSpansPerCandidate));
            if (source.MaxSpanCharactersPerCandidate < 1) throw new ArgumentOutOfRangeException(nameof(source.MaxSpanCharactersPerCandidate));
            if (source.MaxResultMetadataCharactersPerCandidate < 1) throw new ArgumentOutOfRangeException(nameof(source.MaxResultMetadataCharactersPerCandidate));
            if (source.MaxProviderDiagnosticsPerCandidate < 1) throw new ArgumentOutOfRangeException(nameof(source.MaxProviderDiagnosticsPerCandidate));
            if (source.MaxProviderDiagnosticCharactersPerCandidate < 1) throw new ArgumentOutOfRangeException(nameof(source.MaxProviderDiagnosticCharactersPerCandidate));
            if (source.MaxProviderDiagnosticAttributesPerCandidate < 1) throw new ArgumentOutOfRangeException(nameof(source.MaxProviderDiagnosticAttributesPerCandidate));
            if (source.MaxProviderDiagnosticAttributeCharactersPerCandidate < 1) throw new ArgumentOutOfRangeException(nameof(source.MaxProviderDiagnosticAttributeCharactersPerCandidate));
            return new ExecutionOptionsSnapshot {
                Language = string.IsNullOrWhiteSpace(source.Language) ? null : source.Language!.Trim(),
                MaxCandidates = source.MaxCandidates,
                MaxInputBytesPerCandidate = source.MaxInputBytesPerCandidate,
                MaxTotalInputBytes = source.MaxTotalInputBytes,
                MaxDegreeOfParallelism = source.MaxDegreeOfParallelism,
                CandidateTimeout = source.CandidateTimeout,
                MaxRecognizedCharactersPerCandidate = source.MaxRecognizedCharactersPerCandidate,
                MaxSpansPerCandidate = source.MaxSpansPerCandidate,
                MaxSpanCharactersPerCandidate = source.MaxSpanCharactersPerCandidate,
                MaxResultMetadataCharactersPerCandidate = source.MaxResultMetadataCharactersPerCandidate,
                MaxProviderDiagnosticsPerCandidate = source.MaxProviderDiagnosticsPerCandidate,
                MaxProviderDiagnosticCharactersPerCandidate = source.MaxProviderDiagnosticCharactersPerCandidate,
                MaxProviderDiagnosticAttributesPerCandidate = source.MaxProviderDiagnosticAttributesPerCandidate,
                MaxProviderDiagnosticAttributeCharactersPerCandidate = source.MaxProviderDiagnosticAttributeCharactersPerCandidate,
                ContinueOnError = source.ContinueOnError,
                RequirePayloadHashMatch = source.RequirePayloadHashMatch,
                ProviderOptions = source.ProviderOptions == null
                    ? new Dictionary<string, string>(StringComparer.Ordinal)
                    : source.ProviderOptions.ToDictionary(static pair => pair.Key, static pair => pair.Value, StringComparer.Ordinal),
                EnrichmentOptions = CloneEnrichmentOptions(source.EnrichmentOptions)
            };
        }

        private static OfficeDocumentOcrEnrichmentOptions CloneEnrichmentOptions(OfficeDocumentOcrEnrichmentOptions? source) {
            source ??= new OfficeDocumentOcrEnrichmentOptions();
            return new OfficeDocumentOcrEnrichmentOptions {
                RemoveResolvedCandidates = source.RemoveResolvedCandidates,
                RemoveResolvedOcrNeededDiagnostics = source.RemoveResolvedOcrNeededDiagnostics,
                AppendRecognizedTextToMarkdown = source.AppendRecognizedTextToMarkdown,
                BlockKind = source.BlockKind
            };
        }
    }
}
