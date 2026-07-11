namespace OfficeIMO.OpenDocument;

/// <summary>Feature inspection result for an opened document.</summary>
public sealed class OdfFeatureReport {
    internal OdfFeatureReport(IReadOnlyList<OdfFeatureFinding> findings) {
        Findings = findings;
    }

    /// <summary>Detected features and support levels.</summary>
    public IReadOnlyList<OdfFeatureFinding> Findings { get; }
}
