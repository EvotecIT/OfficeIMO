namespace OfficeIMO.OpenDocument;

/// <summary>Feature inspection result for an opened document.</summary>
public sealed class OdfFeatureReport {
    internal OdfFeatureReport(IReadOnlyList<OdfFeatureFinding> findings, IReadOnlyList<OdfFeatureDiagnostic>? diagnostics = null) {
        Findings = findings;
        Diagnostics = diagnostics ?? Array.Empty<OdfFeatureDiagnostic>();
    }

    /// <summary>Detected features and support levels.</summary>
    public IReadOnlyList<OdfFeatureFinding> Findings { get; }
    /// <summary>Inspection failures that prevented one or more package parts from being classified.</summary>
    public IReadOnlyList<OdfFeatureDiagnostic> Diagnostics { get; }
    /// <summary>True when every XML package part was inspected successfully.</summary>
    public bool IsComplete => Diagnostics.Count == 0;
}
