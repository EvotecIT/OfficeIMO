namespace OfficeIMO.Reader;

public static partial class OfficeDocumentOcrEnrichmentExtensions {
    internal static string? RecognitionIdentity(string? value) => value == null || value.Length > 256 || value.Any(char.IsControl) ? null : value;

    private static OfficeDocumentRecognitionEvidence BuildRecognitionEvidence(OfficeDocumentOcrTextResult result) {
        OfficeDocumentRecognitionEvidence? evidence = result.Recognition;
        double? confidence = result.Confidence;
        if (confidence.HasValue && (double.IsNaN(confidence.Value) || double.IsInfinity(confidence.Value) || confidence < 0 || confidence > 1)) confidence = null;
        return new OfficeDocumentRecognitionEvidence(RecognitionIdentity(result.Provider), RecognitionIdentity(result.Model), RecognitionIdentity(result.Language), confidence,
            evidence?.WordCount, evidence?.LowConfidenceWordCount, evidence?.UnknownConfidenceWordCount,
            evidence?.ConfidenceChecksPassed, evidence?.CompletedAttempts, evidence?.HasDisagreement,
            evidence?.ComparisonIncomplete, evidence?.ReviewRecommended, evidence?.MinimumWordConfidence, evidence?.MaximumUncertainWordFraction);
    }
}
