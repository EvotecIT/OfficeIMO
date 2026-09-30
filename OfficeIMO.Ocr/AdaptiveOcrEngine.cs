using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Ocr;

/// <summary>Opt-in, bounded retries of weak OCR evidence using caller-configured recognition variants.</summary>
/// <remarks>
/// Variants receive identical raster bytes, metadata, and coordinate frames. Configure segmentation or language
/// in each provider; raster preparation remains in the format/scanning owner. Retries cannot silently approve text.
/// </remarks>
public sealed partial class AdaptiveOcrEngine : IOcrEngine {
    private readonly OcrEngineExecution[] _engines;
    private readonly string[] _names;
    private readonly OcrReviewPolicy _policy;
    private readonly TimeSpan _timeout;
    private readonly OcrResultCaptureLimits _capture;
    private readonly int _maximumTextCharacters;
    private readonly long _maximumInputBytes;

    /// <summary>Creates up to four ordered variants with a single total execution budget.</summary>
    /// <param name="id">Stable identifier for this configured policy.</param>
    /// <param name="attempts">First attempt is the baseline. Later variants run only when its checks fail.</param>
    /// <param name="reviewPolicy">Measurable checks; defaults are starting settings, not calibrated accuracy claims.</param>
    /// <param name="timeout">Total budget including every gate wait, result capture, and retry. Defaults to one minute.</param>
    /// <param name="maximumSpans">Maximum retained spans per attempt.</param>
    /// <param name="maximumTextCharacters">Maximum text characters per attempt, including span text.</param>
    /// <param name="maximumInputBytes">Maximum encoded raster bytes; each attempt receives an isolated copy.</param>
    public AdaptiveOcrEngine(string id, IReadOnlyList<OcrRecognitionAttempt> attempts,
        OcrReviewPolicy? reviewPolicy = null, TimeSpan? timeout = null, int maximumSpans = 100_000,
        int maximumTextCharacters = 1_000_000, long maximumInputBytes = 25L * 1024 * 1024) {
        Id = OcrEngineRunner.ValidateEngineId(id, nameof(id));
        if (attempts == null) throw new ArgumentNullException(nameof(attempts));
        if (attempts.Count < 1 || attempts.Count > 4) throw new ArgumentOutOfRangeException(nameof(attempts));
        _timeout = timeout ?? TimeSpan.FromMinutes(1);
        if (_timeout <= TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(timeout));
        if (maximumSpans < 1) throw new ArgumentOutOfRangeException(nameof(maximumSpans));
        if (maximumTextCharacters < 1) throw new ArgumentOutOfRangeException(nameof(maximumTextCharacters));
        if (maximumInputBytes < 1) throw new ArgumentOutOfRangeException(nameof(maximumInputBytes));
        _engines = new OcrEngineExecution[attempts.Count]; _names = new string[attempts.Count];
        for (int index = 0; index < attempts.Count; index++) {
            OcrRecognitionAttempt item = attempts[index] ?? throw new ArgumentException("An attempt cannot be null.", nameof(attempts));
            _names[index] = item.Name; _engines[index] = OcrEngineRunner.CreateExecution(item.Engine);
        }
        if (_names.Distinct(StringComparer.Ordinal).Count() != _names.Length) throw new ArgumentException("Attempt names must be unique.", nameof(attempts));
        _policy = reviewPolicy ?? new OcrReviewPolicy();
        _capture = new OcrResultCaptureLimits(maximumSpans, 128, 512);
        _maximumTextCharacters = maximumTextCharacters; _maximumInputBytes = maximumInputBytes;
    }

    /// <inheritdoc />
    public string Id { get; }

    /// <inheritdoc />
    public OcrEngineCapabilities Capabilities {
        get {
            OcrEngineCapabilities capabilities = _engines[0].Capabilities;
            // Shared runner gates serialize each underlying provider. Per-call adaptive state is independent.
            capabilities.SupportsConcurrentRequests = true;
            return capabilities;
        }
    }

    /// <inheritdoc />
    public async Task<OcrResult> RecognizeAsync(OcrRequest request, CancellationToken cancellationToken = default) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        if (request.Operation == OcrOperation.DetectOrientation) {
            ValidateInput(request, cancellationToken);
            Stopwatch clock = Stopwatch.StartNew();
            OcrRequest snapshot = CopyRequest(request);
            TimeSpan remaining = _timeout - clock.Elapsed;
            if (remaining <= TimeSpan.Zero) throw new OcrEngineTimeoutException(Id, _timeout, providerCallStarted: false);
            return await _engines[0].RecognizeAsync(snapshot, remaining, _capture, cancellationToken).ConfigureAwait(false);
        }
        return (await RecognizeWithReviewAsync(request, cancellationToken).ConfigureAwait(false)).Result;
    }

    /// <summary>Recognizes text and returns content-free attempt, uncertainty, disagreement, and selection evidence.</summary>
    public async Task<AdaptiveOcrResult> RecognizeWithReviewAsync(OcrRequest request, CancellationToken cancellationToken = default) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        if (request.Operation != OcrOperation.RecognizeText) throw new ArgumentOutOfRangeException(nameof(request.Operation));
        ValidateInput(request, cancellationToken);
        Stopwatch clock = Stopwatch.StartNew();
        OcrRequest snapshot = CopyRequest(request);
        var summaries = new List<OcrAttemptAssessment>();
        OcrResult? selected = null;
        OcrQualityAssessment? selectedQuality = null;
        int selectedIndex = 0, baselineWords = 0;
        string? baselineText = null;
        bool disagreement = false, incomplete = false;
        for (int index = 0; index < _engines.Length; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            TimeSpan remaining = _timeout - clock.Elapsed;
            if (remaining <= TimeSpan.Zero) { incomplete = true; break; }
            Stopwatch attemptClock = Stopwatch.StartNew();
            if (!_engines[index].Capabilities.SupportsMediaType(snapshot.MediaType)) {
                if (index == 0) throw new NotSupportedException("The baseline OCR variant does not support this media type.");
                summaries.Add(new OcrAttemptAssessment(_names[index], null, "unsupported-media", attemptClock.Elapsed));
                incomplete = true; continue;
            }
            OcrResult candidate;
            OcrQualityAssessment quality;
            string normalized;
            try {
                OcrRequest attemptRequest = CopyRequest(snapshot);
                remaining = _timeout - clock.Elapsed;
                if (remaining <= TimeSpan.Zero) throw new OcrEngineTimeoutException(Id, _timeout, providerCallStarted: false);
                candidate = await _engines[index].RecognizeAsync(attemptRequest, remaining, _capture, cancellationToken).ConfigureAwait(false);
                CheckTextBudget(candidate);
                cancellationToken.ThrowIfCancellationRequested();
                quality = _policy.Assess(candidate);
                try { normalized = Normalize(candidate.Text, cancellationToken); }
                catch (ArgumentException) { throw new OcrEngineExecutionException(OcrEngineFailureKind.InvalidResult); }
                if (clock.Elapsed >= _timeout) throw new OcrEngineTimeoutException(Id, _timeout, providerCallStarted: true);
            } catch (Exception error) when (index > 0 && (error is OcrEngineExecutionException || error is OcrEngineTimeoutException)) {
                cancellationToken.ThrowIfCancellationRequested();
                summaries.Add(new OcrAttemptAssessment(_names[index], null,
                    error is OcrEngineTimeoutException ? "timed-out" : "failed", attemptClock.Elapsed));
                incomplete = true;
                // A timed-out provider may still be running; do not start more work after budget exhaustion.
                if (error is OcrEngineTimeoutException) break;
                continue;
            }
            cancellationToken.ThrowIfCancellationRequested();
            summaries.Add(new OcrAttemptAssessment(_names[index], quality, "completed", attemptClock.Elapsed));
            if (index == 0) {
                selected = candidate; selectedQuality = quality; baselineWords = quality.WordCount; baselineText = normalized;
            } else {
                disagreement |= !string.Equals(baselineText, normalized, StringComparison.Ordinal);
                // A retry must retain baseline coverage and improve word-level uncertainty. Overall confidence
                // never ranks candidates, and disagreement remains a mandatory review signal.
                bool retained = quality.WordCount >= baselineWords * _policy.MinimumRetainedWordRatio;
                bool usable = quality.WordCount >= _policy.MinimumWordCount && !string.IsNullOrWhiteSpace(candidate.Text) &&
                    !quality.HasWarningsOrErrors && !quality.HasOmittedSpans;
                if (retained && usable && (selectedQuality!.HasWarningsOrErrors || selectedQuality.HasOmittedSpans ||
                    selectedQuality.WordCount < _policy.MinimumWordCount || quality.UncertainWordFraction < selectedQuality.UncertainWordFraction)) {
                    selected = candidate; selectedQuality = quality; selectedIndex = summaries.Count - 1;
                }
            }
            if (selectedQuality!.MeetsThresholds) break;
        }
        cancellationToken.ThrowIfCancellationRequested();
        if (selected == null || selectedQuality == null) throw new OcrEngineTimeoutException(Id, _timeout, providerCallStarted: false);
        var result = new AdaptiveOcrResult(selected, selectedIndex, summaries.AsReadOnly(), disagreement, incomplete, selectedQuality);
        AddReviewDiagnostic(result);
        return result;
    }
}
