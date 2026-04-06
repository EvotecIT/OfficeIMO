namespace OfficeIMO.Markdown;

/// <summary>
/// Lightweight diagnostic emitted for each document transform application.
/// </summary>
public sealed class MarkdownDocumentTransformDiagnostic {
    /// <summary>
    /// Known source stage that invoked the transform pipeline.
    /// </summary>
    public MarkdownDocumentTransformSource Source { get; set; }

    /// <summary>
    /// Transform type name.
    /// </summary>
    public string TransformName { get; set; } = string.Empty;

    /// <summary>
    /// Number of top-level blocks before the transform ran.
    /// </summary>
    public int BlockCountBefore { get; set; }

    /// <summary>
    /// Number of top-level blocks after the transform ran.
    /// </summary>
    public int BlockCountAfter { get; set; }

    /// <summary>
    /// Whether the transform returned a different document instance.
    /// </summary>
    public bool ReplacedDocument { get; set; }

    /// <summary>
    /// First top-level block index affected before the transform ran.
    /// </summary>
    public int ChangedBlockStartBefore { get; set; }

    /// <summary>
    /// Number of contiguous top-level blocks affected before the transform ran.
    /// </summary>
    public int ChangedBlockCountBefore { get; set; }

    /// <summary>
    /// First top-level block index affected after the transform ran.
    /// </summary>
    public int ChangedBlockStartAfter { get; set; }

    /// <summary>
    /// Number of contiguous top-level blocks affected after the transform ran.
    /// </summary>
    public int ChangedBlockCountAfter { get; set; }

    /// <summary>
    /// Aggregate source span of the affected input blocks when original syntax spans are available.
    /// </summary>
    public MarkdownSourceSpan? AffectedSourceSpan { get; set; }

    /// <summary>
    /// Deepest original block-like syntax path overlapping the affected source span when available.
    /// </summary>
    public string? AffectedOriginalBlockPath { get; set; }

    /// <summary>
    /// Source span of the original block-like syntax node identified by <see cref="AffectedOriginalBlockPath"/>.
    /// </summary>
    public MarkdownSourceSpan? AffectedOriginalBlockSpan { get; set; }

    /// <summary>
    /// Deepest final block-like syntax path overlapping the affected source span when available.
    /// </summary>
    public string? AffectedFinalBlockPath { get; set; }

    /// <summary>
    /// Source span of the final block-like syntax node identified by <see cref="AffectedFinalBlockPath"/>.
    /// </summary>
    public MarkdownSourceSpan? AffectedFinalBlockSpan { get; set; }
}
