namespace OfficeIMO.Ocr;

/// <summary>Controls which caller-configured recognition variants run within the shared deadline.</summary>
public enum OcrRetryMode {
    /// <summary>Stop after an attempt meets the configured confidence checks.</summary>
    WhenUncertain,
    /// <summary>Run every configured variant, including when the baseline meets confidence checks.</summary>
    CompareAll
}

/// <summary>Recognition review state; no value constitutes human approval or semantic correctness.</summary>
public enum OcrReviewStatus {
    /// <summary>Confidence checks passed, but no second completed variant corroborated the text.</summary>
    Unassessed,
    /// <summary>At least two completed variants agreed and the selected evidence passed configured checks.</summary>
    ChecksPassed,
    /// <summary>Uncertainty, disagreement or an incomplete comparison requires attention.</summary>
    ReviewRecommended
}

/// <summary>Immutable evidence from adaptive recognition, preserved by the shared runner.</summary>
public sealed class OcrReviewEvidence {
    internal OcrReviewEvidence(OcrQualityAssessment quality, int completedAttempts, bool disagreement, bool incomplete) {
        Quality = quality; CompletedAttempts = completedAttempts; HasDisagreement = disagreement; ComparisonIncomplete = incomplete;
    }
    /// <summary>Word confidence checks before adaptive diagnostics were attached.</summary>
    public OcrQualityAssessment Quality { get; }
    /// <summary>Number of completed recognition variants, including the baseline.</summary>
    public int CompletedAttempts { get; }
    /// <summary>Completed variants returned differing normalized text.</summary>
    public bool HasDisagreement { get; }
    /// <summary>A configured comparison failed, exceeded its budget, or lost retained evidence.</summary>
    public bool ComparisonIncomplete { get; }
    /// <summary>Check status, deliberately separate from approval and correctness.</summary>
    public OcrReviewStatus Status => !Quality.MeetsThresholds || HasDisagreement || ComparisonIncomplete
        ? OcrReviewStatus.ReviewRecommended : CompletedAttempts < 2 ? OcrReviewStatus.Unassessed : OcrReviewStatus.ChecksPassed;
}
