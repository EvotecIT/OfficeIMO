using System;

namespace OfficeIMO.Reader;

/// <summary>Immutable OCR provenance attached to recognized text. Missing values mean unknown, never approval.</summary>
/// <remarks>Counts and confidence describe recognition checks, not the probability that text is correct.</remarks>
public sealed class OfficeDocumentRecognitionEvidence {
    /// <summary>Captures bounded provider identity and available recognition/review measurements.</summary>
    public OfficeDocumentRecognitionEvidence(string? provider = null, string? model = null, string? language = null,
        double? confidence = null, int? wordCount = null, int? lowConfidenceWordCount = null,
        int? unknownConfidenceWordCount = null, bool? confidenceChecksPassed = null,
        int? completedAttempts = null, bool? hasDisagreement = null, bool? comparisonIncomplete = null,
        bool? reviewRecommended = null, double? minimumWordConfidence = null, double? maximumUncertainWordFraction = null) {
        Provider = Identity(provider, nameof(provider)); Model = Identity(model, nameof(model)); Language = Identity(language, nameof(language));
        if (confidence.HasValue && (double.IsNaN(confidence.Value) || double.IsInfinity(confidence.Value) || confidence < 0 || confidence > 1))
            throw new ArgumentOutOfRangeException(nameof(confidence));
        if (wordCount < 0 || lowConfidenceWordCount < 0 || unknownConfidenceWordCount < 0 || completedAttempts < 0
            || (wordCount.HasValue && (long)(lowConfidenceWordCount ?? 0) + (unknownConfidenceWordCount ?? 0) > wordCount.Value))
            throw new ArgumentOutOfRangeException(nameof(wordCount));
        Confidence = confidence; WordCount = wordCount; LowConfidenceWordCount = lowConfidenceWordCount;
        UnknownConfidenceWordCount = unknownConfidenceWordCount; ConfidenceChecksPassed = confidenceChecksPassed;
        CompletedAttempts = completedAttempts; HasDisagreement = hasDisagreement; ComparisonIncomplete = comparisonIncomplete;
        ReviewRecommended = reviewRecommended;
        MinimumWordConfidence = Unit(minimumWordConfidence, nameof(minimumWordConfidence));
        MaximumUncertainWordFraction = Unit(maximumUncertainWordFraction, nameof(maximumUncertainWordFraction));
    }
    /// <summary>Non-secret provider identifier, when available.</summary>
    public string? Provider { get; }
    /// <summary>Recognition model or engine version, when available.</summary>
    public string? Model { get; }
    /// <summary>Provider language expression, when available.</summary>
    public string? Language { get; }
    /// <summary>Aggregate provider confidence, not a correctness probability.</summary>
    public double? Confidence { get; }
    /// <summary>Word spans assessed before any downstream retention loss.</summary>
    public int? WordCount { get; }
    /// <summary>Assessed words below the configured confidence threshold.</summary>
    public int? LowConfidenceWordCount { get; }
    /// <summary>Assessed words with unknown or invalid confidence.</summary>
    public int? UnknownConfidenceWordCount { get; }
    /// <summary>Whether configured confidence checks passed; null when not assessed.</summary>
    public bool? ConfidenceChecksPassed { get; }
    /// <summary>Completed recognition variants, when recorded.</summary>
    public int? CompletedAttempts { get; }
    /// <summary>Whether completed variants disagreed; null when not compared.</summary>
    public bool? HasDisagreement { get; }
    /// <summary>Whether comparison or evidence retention was incomplete.</summary>
    public bool? ComparisonIncomplete { get; }
    /// <summary>Recorded need for review; false means checks passed, never human approval.</summary>
    public bool? ReviewRecommended { get; }

    /// <summary>Threshold used to classify low-confidence words, when known.</summary>
    public double? MinimumWordConfidence { get; }
    /// <summary>Maximum uncertain fraction allowed by the recorded checks, when known.</summary>
    public double? MaximumUncertainWordFraction { get; }

    private static double? Unit(double? value, string name) {
        if (value.HasValue && (double.IsNaN(value.Value) || double.IsInfinity(value.Value) || value < 0 || value > 1))
            throw new ArgumentOutOfRangeException(name);
        return value;
    }

    private static string? Identity(string? value, string name) {
        if (value == null) return null;
        if (value.Length > 256) throw new ArgumentOutOfRangeException(name);
        foreach (char character in value) if (char.IsControl(character)) throw new ArgumentException("Recognition identity cannot contain control characters.", name);
        return value;
    }
}
