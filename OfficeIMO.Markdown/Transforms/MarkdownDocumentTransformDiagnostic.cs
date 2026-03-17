namespace OfficeIMO.Markdown;

/// <summary>
/// Lightweight diagnostic emitted for each document transform application.
/// </summary>
public sealed class MarkdownDocumentTransformDiagnostic {
    /// <summary>
    /// Known source stage that invoked the transform pipeline.
    /// </summary>
    public MarkdownDocumentTransformSource Source { get; init; }

    /// <summary>
    /// Transform type name.
    /// </summary>
    public string TransformName { get; init; } = string.Empty;

    /// <summary>
    /// Number of top-level blocks before the transform ran.
    /// </summary>
    public int BlockCountBefore { get; init; }

    /// <summary>
    /// Number of top-level blocks after the transform ran.
    /// </summary>
    public int BlockCountAfter { get; init; }

    /// <summary>
    /// Whether the transform returned a different document instance.
    /// </summary>
    public bool ReplacedDocument { get; init; }

    /// <summary>
    /// First top-level block index affected before the transform ran.
    /// </summary>
    public int ChangedBlockStartBefore { get; init; }

    /// <summary>
    /// Number of contiguous top-level blocks affected before the transform ran.
    /// </summary>
    public int ChangedBlockCountBefore { get; init; }

    /// <summary>
    /// First top-level block index affected after the transform ran.
    /// </summary>
    public int ChangedBlockStartAfter { get; init; }

    /// <summary>
    /// Number of contiguous top-level blocks affected after the transform ran.
    /// </summary>
    public int ChangedBlockCountAfter { get; init; }

    /// <summary>
    /// Aggregate source span of the affected input blocks when original syntax spans are available.
    /// </summary>
    public MarkdownSourceSpan? AffectedSourceSpan { get; init; }
}
