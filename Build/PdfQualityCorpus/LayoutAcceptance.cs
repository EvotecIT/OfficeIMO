using System.Text.Json;

namespace OfficeIMO.PdfQualityCorpus;

/// <summary>Predeclared layout-corpus limits, independent of measured product output.</summary>
internal sealed record LayoutAcceptance(
    double MaximumCharacterErrorRate,
    double MaximumWordErrorRate,
    bool RequireCompleteReadingOrder,
    bool RequireExactTables) {
    internal static LayoutAcceptance? Read(JsonElement fixture) {
        if (!fixture.TryGetProperty("acceptance", out JsonElement value)) return null;
        double cer = value.GetProperty("maximumCharacterErrorRate").GetDouble();
        double wer = value.GetProperty("maximumWordErrorRate").GetDouble();
        if (!double.IsFinite(cer) || !double.IsFinite(wer) || cer < 0 || wer < 0)
            throw new InvalidDataException("Layout acceptance error-rate limits must be finite and nonnegative.");
        return new LayoutAcceptance(cer, wer, value.GetProperty("requireCompleteReadingOrder").GetBoolean(),
            value.GetProperty("requireExactTables").GetBoolean());
    }

    internal bool IsSatisfied(ScanTextAccuracy accuracy, int presentSegments, int expectedSegments,
        int correctPairs, int expectedPairs, int exactTableRows, int expectedTableRows, int detectedTables) =>
        accuracy.CharacterErrorRate <= MaximumCharacterErrorRate &&
        accuracy.WordErrorRate <= MaximumWordErrorRate &&
        (!RequireCompleteReadingOrder || (presentSegments == expectedSegments && correctPairs == expectedPairs)) &&
        (!RequireExactTables || (exactTableRows == expectedTableRows &&
            detectedTables == (expectedTableRows == 0 ? 0 : 1)));
}
