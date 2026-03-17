namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Lightweight diagnostic emitted for each renderer pre-parse processing step.
/// </summary>
public sealed class MarkdownRendererPreProcessorDiagnostic {
    /// <summary>
    /// Pre-parse stage that produced this diagnostic.
    /// </summary>
    public MarkdownRendererPreProcessorStage Stage { get; set; }

    /// <summary>
    /// Processor name for custom pre-processors.
    /// </summary>
    public string ProcessorName { get; set; } = string.Empty;

    /// <summary>
    /// Input length before the stage ran.
    /// </summary>
    public int LengthBefore { get; set; }

    /// <summary>
    /// Input length after the stage ran.
    /// </summary>
    public int LengthAfter { get; set; }

    /// <summary>
    /// Whether the stage changed the markdown text.
    /// </summary>
    public bool Changed { get; set; }
}

/// <summary>
/// Known renderer pre-parse stages that can emit diagnostics.
/// </summary>
public enum MarkdownRendererPreProcessorStage {
    /// <summary>Renderer escaped newline normalization.</summary>
    EscapedNewlineNormalization = 0,
    /// <summary>Shared markdown input normalization before parsing.</summary>
    InputNormalization = 1,
    /// <summary>Custom renderer markdown pre-processor.</summary>
    CustomPreProcessor = 2
}
