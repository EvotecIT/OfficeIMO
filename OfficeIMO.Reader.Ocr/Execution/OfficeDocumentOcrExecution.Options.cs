using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Reader;

public static partial class OfficeDocumentOcrExecutionExtensions {
    private sealed class ExecutionOptionsSnapshot {
        private ExecutionOptionsSnapshot() { }

        internal string? Language { get; private set; }
        internal int MaxCandidates { get; private set; }
        internal long MaxInputBytesPerCandidate { get; private set; }
        internal long MaxTotalInputBytes { get; private set; }
        internal int MaxDegreeOfParallelism { get; private set; }
        internal TimeSpan CandidateTimeout { get; private set; }
        internal TimeSpan TotalTimeout { get; private set; }
        internal int MaxDocuments { get; private set; }
        internal int MaxNestedDepth { get; private set; }
        internal int MaxTotalRecognizedCharacters { get; private set; }
        internal int MaxTotalSpans { get; private set; }
        internal int MaxTotalSpanCharacters { get; private set; }
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
            if (source.TotalTimeout <= TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(source.TotalTimeout));
            if (source.MaxDocuments < 1) throw new ArgumentOutOfRangeException(nameof(source.MaxDocuments));
            if (source.MaxNestedDepth < 0 || source.MaxNestedDepth > 64) throw new ArgumentOutOfRangeException(nameof(source.MaxNestedDepth));
            if (source.MaxTotalRecognizedCharacters < 1) throw new ArgumentOutOfRangeException(nameof(source.MaxTotalRecognizedCharacters));
            if (source.MaxTotalSpans < 1) throw new ArgumentOutOfRangeException(nameof(source.MaxTotalSpans));
            if (source.MaxTotalSpanCharacters < 1) throw new ArgumentOutOfRangeException(nameof(source.MaxTotalSpanCharacters));
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
                TotalTimeout = source.TotalTimeout,
                MaxDocuments = source.MaxDocuments,
                MaxNestedDepth = source.MaxNestedDepth,
                MaxTotalRecognizedCharacters = source.MaxTotalRecognizedCharacters,
                MaxTotalSpans = source.MaxTotalSpans,
                MaxTotalSpanCharacters = source.MaxTotalSpanCharacters,
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
