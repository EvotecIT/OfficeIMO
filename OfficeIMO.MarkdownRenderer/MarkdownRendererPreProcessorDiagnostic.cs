namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Lightweight diagnostic emitted for each renderer pre-parse processing step.
/// </summary>
public sealed class MarkdownRendererPreProcessorDiagnostic {
    /// <summary>
    /// Pre-parse stage that produced this diagnostic.
    /// </summary>
    public MarkdownRendererPreProcessorStage Stage { get; init; }

    /// <summary>
    /// Processor name for custom pre-processors.
    /// </summary>
    public string ProcessorName { get; init; } = string.Empty;

    /// <summary>
    /// Input length before the stage ran.
    /// </summary>
    public int LengthBefore { get; init; }

    /// <summary>
    /// Input length after the stage ran.
    /// </summary>
    public int LengthAfter { get; init; }

    /// <summary>
    /// Whether the stage changed the markdown text.
    /// </summary>
    public bool Changed { get; init; }
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
