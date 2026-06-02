namespace OfficeIMO.Markdown;

/// <summary>
/// Named input-normalization presets for common markdown ingestion scenarios.
/// </summary>
public enum MarkdownInputNormalizationPreset {
    /// <summary>No preset behavior.</summary>
    None = 0,
    /// <summary>
     /// Conservative transcript repair preset aligned with the explicit IntelligenceX transcript contract.
     /// </summary>
    IntelligenceXTranscript = 1,
    /// <summary>
    /// Broader IntelligenceX transcript repair preset for aggressively malformed transcript content.
    /// </summary>
    IntelligenceXTranscriptStrict = 2,
    /// <summary>
     /// Conservative documentation import preset that avoids transcript-specific boundary rewrites.
     /// </summary>
    DocsLoose = 3,
}
