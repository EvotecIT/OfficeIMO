using System;
using System.Collections.Generic;

namespace OfficeIMO.Ocr;

/// <summary>One caller-configured recognition variant. All variants receive the same raster and coordinate frame.</summary>
public sealed class OcrRecognitionAttempt {
    /// <summary>Creates a named variant such as <c>psm-6</c>. Engines remain caller-owned.</summary>
    public OcrRecognitionAttempt(string name, IOcrEngine engine) {
        if (string.IsNullOrEmpty(name) || name.Length > 64) throw new ArgumentException("An attempt name must contain 1 through 64 ASCII letters, digits, or hyphens.", nameof(name));
        foreach (char character in name) {
            if (!(character >= 'a' && character <= 'z' || character >= 'A' && character <= 'Z' ||
                character >= '0' && character <= '9' || character == '-')) throw new ArgumentException("Invalid attempt name.", nameof(name));
        }
        Name = name; Engine = engine ?? throw new ArgumentNullException(nameof(engine));
    }
    /// <summary>Content-free variant identifier.</summary>
    public string Name { get; }
    /// <summary>Configured engine. The adaptive engine captures its identity and capabilities at construction.</summary>
    public IOcrEngine Engine { get; }
}

/// <summary>Content-free evidence for one attempted recognition variant.</summary>
public sealed class OcrAttemptAssessment {
    internal OcrAttemptAssessment(string name, OcrQualityAssessment? quality, string outcome, TimeSpan elapsed) {
        Name = name; Quality = quality; Outcome = outcome; Elapsed = elapsed;
    }
    /// <summary>Configured variant name.</summary>
    public string Name { get; }
    /// <summary>Quality checks, absent when the attempt failed.</summary>
    public OcrQualityAssessment? Quality { get; }
    /// <summary>Stable outcome: completed, failed, timed-out, or unsupported-media.</summary>
    public string Outcome { get; }
    /// <summary>Elapsed attempt time including gate waits.</summary>
    public TimeSpan Elapsed { get; }
}

/// <summary>Adaptive recognition result and review evidence. Text remains proposed until reviewed.</summary>
public sealed class AdaptiveOcrResult {
    internal AdaptiveOcrResult(OcrResult result, int selected, IReadOnlyList<OcrAttemptAssessment> attempts,
        bool disagreement, bool interrupted, OcrQualityAssessment quality) {
        Result = result; SelectedAttempt = selected; Attempts = attempts;
        HasDisagreement = disagreement; RetryIncomplete = interrupted; Quality = quality;
        int completed = 0;
        foreach (OcrAttemptAssessment attempt in attempts) if (attempt.Quality != null) completed++;
        Review = new OcrReviewEvidence(quality, completed, disagreement, interrupted);
        result.Review = Review;
    }
    /// <summary>Owned output from the selected variant; geometry belongs to the original request raster.</summary>
    public OcrResult Result { get; }
    /// <summary>Zero-based index within the attempted variants.</summary>
    public int SelectedAttempt { get; }
    /// <summary>At most four bounded attempt summaries.</summary>
    public IReadOnlyList<OcrAttemptAssessment> Attempts { get; }
    /// <summary>Whether completed attempts disagreed after NFC and whitespace normalization. Case/punctuation remain significant.</summary>
    public bool HasDisagreement { get; }
    /// <summary>Whether an attempt lost evidence, failed, timed out, was unsupported, or could not start within the budget.</summary>
    public bool RetryIncomplete { get; }
    /// <summary>Selected output's checks before adaptive diagnostics are added.</summary>
    public OcrQualityAssessment Quality { get; }
    /// <summary>Immutable comparison evidence. A single passing attempt remains unassessed.</summary>
    public OcrReviewEvidence Review { get; }
    /// <summary>Whether recognition lacks corroboration or its checks recommend attention.</summary>
    public bool ReviewRecommended => Review.Status != OcrReviewStatus.ChecksPassed;
}
