using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Reader;

/// <summary>Controls bounded execution of an optional OCR engine over a document read result.</summary>
public sealed class OfficeDocumentOcrExecutionOptions {
    /// <summary>Language or provider-specific language expression requested for every candidate.</summary>
    public string? Language { get; set; }

    /// <summary>Maximum candidates considered in one execution. Defaults to 100.</summary>
    public int MaxCandidates { get; set; } = 100;

    /// <summary>Maximum payload size passed to the engine for one candidate. Defaults to 25 MiB.</summary>
    public long MaxInputBytesPerCandidate { get; set; } = 25L * 1024L * 1024L;

    /// <summary>Maximum combined payload size passed to the engine. Defaults to 100 MiB.</summary>
    public long MaxTotalInputBytes { get; set; } = 100L * 1024L * 1024L;

    /// <summary>Maximum concurrent engine calls. Defaults to 1 and is forced to 1 for engines that do not advertise concurrency.</summary>
    public int MaxDegreeOfParallelism { get; set; } = 1;

    /// <summary>Maximum duration allowed for one engine call. Defaults to two minutes.</summary>
    public TimeSpan CandidateTimeout { get; set; } = TimeSpan.FromMinutes(2);

    /// <summary>Maximum recognized text characters accepted from one engine response. Defaults to 1,000,000.</summary>
    public int MaxRecognizedCharactersPerCandidate { get; set; } = 1_000_000;

    /// <summary>Maximum detailed spans accepted from one engine response. Defaults to 100,000.</summary>
    public int MaxSpansPerCandidate { get; set; } = 100_000;

    /// <summary>Whether one provider failure should become a diagnostic while remaining candidates continue.</summary>
    public bool ContinueOnError { get; set; } = true;

    /// <summary>Whether a declared asset payload hash must match before bytes are passed to an engine.</summary>
    public bool RequirePayloadHashMatch { get; set; } = true;

    /// <summary>Provider-specific scalar options copied into every engine request.</summary>
    public IReadOnlyDictionary<string, string> ProviderOptions { get; set; } = new Dictionary<string, string>(StringComparer.Ordinal);

    /// <summary>Controls how recognized text is merged into the shared read result.</summary>
    public OfficeDocumentOcrEnrichmentOptions EnrichmentOptions { get; set; } = new OfficeDocumentOcrEnrichmentOptions();

    /// <summary>Creates an independent execution configuration snapshot.</summary>
    public OfficeDocumentOcrExecutionOptions Clone() {
        OfficeDocumentOcrEnrichmentOptions enrichment = EnrichmentOptions ?? new OfficeDocumentOcrEnrichmentOptions();
        return new OfficeDocumentOcrExecutionOptions {
            Language = Language,
            MaxCandidates = MaxCandidates,
            MaxInputBytesPerCandidate = MaxInputBytesPerCandidate,
            MaxTotalInputBytes = MaxTotalInputBytes,
            MaxDegreeOfParallelism = MaxDegreeOfParallelism,
            CandidateTimeout = CandidateTimeout,
            MaxRecognizedCharactersPerCandidate = MaxRecognizedCharactersPerCandidate,
            MaxSpansPerCandidate = MaxSpansPerCandidate,
            ContinueOnError = ContinueOnError,
            RequirePayloadHashMatch = RequirePayloadHashMatch,
            ProviderOptions = ProviderOptions == null
                ? new Dictionary<string, string>(StringComparer.Ordinal)
                : ProviderOptions.ToDictionary(static pair => pair.Key, static pair => pair.Value, StringComparer.Ordinal),
            EnrichmentOptions = new OfficeDocumentOcrEnrichmentOptions {
                RemoveResolvedCandidates = enrichment.RemoveResolvedCandidates,
                RemoveResolvedOcrNeededDiagnostics = enrichment.RemoveResolvedOcrNeededDiagnostics,
                AppendRecognizedTextToMarkdown = enrichment.AppendRecognizedTextToMarkdown,
                BlockKind = enrichment.BlockKind
            }
        };
    }
}

/// <summary>Observable counters from one bounded OCR execution.</summary>
public sealed class OfficeDocumentOcrExecutionReport {
    /// <summary>Configured engine identifier.</summary>
    public string EngineId { get; set; } = string.Empty;

    /// <summary>Total candidates in the source read result.</summary>
    public int CandidateCount { get; set; }

    /// <summary>Candidates selected before asset and payload validation.</summary>
    public int SelectedCandidateCount { get; set; }

    /// <summary>Candidates whose payloads were passed to the engine.</summary>
    public int AttemptedCandidateCount { get; set; }

    /// <summary>Engine calls that returned non-empty recognized text.</summary>
    public int RecognizedCandidateCount { get; set; }

    /// <summary>Successful engine calls that returned no recognized text.</summary>
    public int EmptyCandidateCount { get; set; }

    /// <summary>Candidates skipped by configured limits, validation checks, or engine capacity held by a timed-out call.</summary>
    public int SkippedCandidateCount { get; set; }

    /// <summary>Engine calls that failed and were converted to diagnostics.</summary>
    public int FailedCandidateCount { get; set; }

    /// <summary>Detailed line spans returned by successful engine calls.</summary>
    public int LineSpanCount { get; set; }

    /// <summary>Detailed word spans returned by successful engine calls.</summary>
    public int WordSpanCount { get; set; }

    /// <summary>Detailed character spans returned by successful engine calls.</summary>
    public int CharacterSpanCount { get; set; }

    /// <summary>Total validated payload bytes passed to the engine.</summary>
    public long InputBytes { get; set; }

    /// <summary>Maximum concurrency actually used for this engine instance.</summary>
    public int EffectiveDegreeOfParallelism { get; set; }
}

/// <summary>Document, detailed recognition output, diagnostics, and counters from optional OCR execution.</summary>
public sealed class OfficeDocumentOcrExecutionResult {
    /// <summary>Document enriched with successfully recognized text.</summary>
    public OfficeDocumentReadResult Document { get; set; } = new OfficeDocumentReadResult();

    /// <summary>Successful engine responses in source candidate order, including empty-text responses.</summary>
    public IReadOnlyList<OfficeDocumentOcrRecognition> Recognitions { get; set; } = Array.Empty<OfficeDocumentOcrRecognition>();

    /// <summary>Execution, provider, validation, and limit diagnostics emitted during this run.</summary>
    public IReadOnlyList<OfficeDocumentDiagnostic> Diagnostics { get; set; } = Array.Empty<OfficeDocumentDiagnostic>();

    /// <summary>Observable execution counters.</summary>
    public OfficeDocumentOcrExecutionReport Report { get; set; } = new OfficeDocumentOcrExecutionReport();
}
