using OfficeIMO.Ocr;

namespace OfficeIMO.Reader;

public static partial class OfficeDocumentOcrExecutionExtensions {
    private static OfficeDocumentRecognitionEvidence CaptureRecognitionEvidence(OcrResult result, bool normalizedWithLoss) {
        OcrReviewEvidence? review = result.Review;
        OcrQualityAssessment quality = review?.Quality ?? new OcrReviewPolicy().Assess(result);
        bool incomplete = normalizedWithLoss || result.OmittedSpanCount > 0 || result.OmittedDiagnosticCount > 0
            || result.Diagnostics.Any(item => item.OmittedAttributeCount > 0 || (item.Severity != OcrDiagnosticSeverity.Info && !item.Code.StartsWith("adaptive-ocr-", StringComparison.Ordinal)));
        return new OfficeDocumentRecognitionEvidence(OfficeDocumentOcrEnrichmentExtensions.RecognitionIdentity(result.Provider),
            OfficeDocumentOcrEnrichmentExtensions.RecognitionIdentity(result.Model),
            OfficeDocumentOcrEnrichmentExtensions.RecognitionIdentity(result.Language), result.Confidence,
            quality.WordCount, quality.LowConfidenceWordCount, quality.UnknownConfidenceWordCount,
            quality.MeetsThresholds && !incomplete, review?.CompletedAttempts,
            review is { CompletedAttempts: > 1 } ? review.HasDisagreement : null,
            incomplete || review?.ComparisonIncomplete == true,
            incomplete || !quality.MeetsThresholds || review?.Status != OcrReviewStatus.ChecksPassed,
            quality.MinimumWordConfidence, quality.MaximumUncertainWordFraction);
    }
}
