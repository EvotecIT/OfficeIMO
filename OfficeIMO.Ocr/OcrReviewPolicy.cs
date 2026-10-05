using System;
using System.Linq;

namespace OfficeIMO.Ocr;

/// <summary>Immutable, measurable checks for deciding whether OCR evidence needs attention.</summary>
/// <remarks>Provider confidence is not a probability of correctness. Passing these checks does not approve text.</remarks>
public sealed class OcrReviewPolicy {
    /// <summary>Creates a policy. Unknown word confidence always counts as uncertain.</summary>
    public OcrReviewPolicy(double minimumWordConfidence = 0.8, double maximumUncertainWordFraction = 0.1,
        int minimumWordCount = 1, double minimumRetainedWordRatio = 0.9)
        : this(OcrRetryMode.WhenUncertain, minimumWordConfidence, maximumUncertainWordFraction, minimumWordCount, minimumRetainedWordRatio) { }

    /// <summary>Creates a policy with an explicit bounded recognition comparison strategy.</summary>
    public OcrReviewPolicy(OcrRetryMode retryMode, double minimumWordConfidence = 0.8,
        double maximumUncertainWordFraction = 0.1, int minimumWordCount = 1, double minimumRetainedWordRatio = 0.9) {
        if (!Enum.IsDefined(typeof(OcrRetryMode), retryMode)) throw new ArgumentOutOfRangeException(nameof(retryMode));
        RetryMode = retryMode;
        ValidateUnit(minimumWordConfidence, nameof(minimumWordConfidence));
        ValidateUnit(maximumUncertainWordFraction, nameof(maximumUncertainWordFraction));
        ValidateUnit(minimumRetainedWordRatio, nameof(minimumRetainedWordRatio));
        if (minimumWordCount < 1) throw new ArgumentOutOfRangeException(nameof(minimumWordCount));
        MinimumWordConfidence = minimumWordConfidence;
        MaximumUncertainWordFraction = maximumUncertainWordFraction;
        MinimumWordCount = minimumWordCount;
        MinimumRetainedWordRatio = minimumRetainedWordRatio;
    }

    /// <summary>Whether to stop after passing confidence checks or compare all configured variants.</summary>
    public OcrRetryMode RetryMode { get; }

    /// <summary>Minimum normalized word score. Equality passes.</summary>
    public double MinimumWordConfidence { get; }
    /// <summary>Maximum fraction of words with low or unknown confidence. Equality passes.</summary>
    public double MaximumUncertainWordFraction { get; }
    /// <summary>Minimum number of nonempty word spans required for scoring.</summary>
    public int MinimumWordCount { get; }
    /// <summary>Minimum word-count ratio relative to the first attempt before a retry may replace it.</summary>
    public double MinimumRetainedWordRatio { get; }

    /// <summary>Assesses a caller-owned result snapshot, without interpreting confidence as correctness.</summary>
    public OcrQualityAssessment Assess(OcrResult result) {
        if (result == null) throw new ArgumentNullException(nameof(result));
        int words = 0, low = 0, unknown = 0;
        foreach (OcrTextSpan span in result.Spans ?? Array.Empty<OcrTextSpan>()) {
            if (span == null || span.Level != OcrTextSpanLevel.Word || string.IsNullOrWhiteSpace(span.Text)) continue;
            words++;
            if (!span.Confidence.HasValue || !ValidUnit(span.Confidence.Value)) unknown++;
            else if (span.Confidence.Value < MinimumWordConfidence) low++;
        }
        bool diagnostics = result.OmittedDiagnosticCount > 0 || (result.Diagnostics ?? Array.Empty<OcrDiagnostic>())
            .Any(item => item == null || item.Severity != OcrDiagnosticSeverity.Info || item.OmittedAttributeCount > 0);
        double fraction = words == 0 ? 1 : (double)(low + unknown) / words;
        bool meets = !string.IsNullOrWhiteSpace(result.Text) && words >= MinimumWordCount &&
            fraction <= MaximumUncertainWordFraction && result.OmittedSpanCount == 0 && !diagnostics;
        return new OcrQualityAssessment(words, low, unknown, fraction, meets, diagnostics, result.OmittedSpanCount > 0, MinimumWordConfidence, MaximumUncertainWordFraction);
    }

    internal static bool HasOmittedEvidence(OcrResult result) => result.OmittedSpanCount > 0 || result.OmittedDiagnosticCount > 0
        || (result.Diagnostics ?? Array.Empty<OcrDiagnostic>()).Any(item => item?.OmittedAttributeCount > 0);

    private static bool ValidUnit(double value) => !double.IsNaN(value) && !double.IsInfinity(value) && value >= 0 && value <= 1;
    private static void ValidateUnit(double value, string name) {
        if (!ValidUnit(value)) throw new ArgumentOutOfRangeException(name);
    }
}

/// <summary>Word-level uncertainty counts for one owned OCR result.</summary>
public sealed class OcrQualityAssessment {
    internal OcrQualityAssessment(int words, int low, int unknown, double fraction, bool meets, bool diagnostics, bool omitted, double minimumWordConfidence, double maximumUncertainWordFraction) {
        MinimumWordConfidence = minimumWordConfidence; MaximumUncertainWordFraction = maximumUncertainWordFraction;
        WordCount = words; LowConfidenceWordCount = low; UnknownConfidenceWordCount = unknown;
        UncertainWordFraction = fraction; MeetsThresholds = meets; HasWarningsOrErrors = diagnostics; HasOmittedSpans = omitted;
    }
    /// <summary>Configured threshold used for these word counts.</summary>
    public double MinimumWordConfidence { get; }
    /// <summary>Configured maximum uncertain fraction used for these checks.</summary>
    public double MaximumUncertainWordFraction { get; }
    /// <summary>Number of nonempty word spans; line/character spans are not counted twice.</summary>
    public int WordCount { get; }
    /// <summary>Words below the configured threshold.</summary>
    public int LowConfidenceWordCount { get; }
    /// <summary>Words whose confidence is absent or invalid.</summary>
    public int UnknownConfidenceWordCount { get; }
    /// <summary>Fraction of low or unknown words; one when no word evidence exists.</summary>
    public double UncertainWordFraction { get; }
    /// <summary>Whether configured checks pass. This is not semantic or human approval.</summary>
    public bool MeetsThresholds { get; }
    /// <summary>Whether warning/error diagnostics exist or diagnostics were omitted.</summary>
    public bool HasWarningsOrErrors { get; }
    /// <summary>Whether span retention omitted evidence.</summary>
    public bool HasOmittedSpans { get; }
}
