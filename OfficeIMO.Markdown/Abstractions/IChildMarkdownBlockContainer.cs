namespace OfficeIMO.Markdown;

/// <summary>
/// Marks a block as a container of nested child blocks.
/// Implement this on custom block types so traversal, source-span binding, syntax-tree generation,
/// rewriters, and descendant queries can walk into nested markdown content.
/// </summary>
public interface IChildMarkdownBlockContainer {
    /// <summary>
    /// Nested markdown child blocks owned by the container.
    /// </summary>
    IReadOnlyList<IMarkdownBlock> ChildBlocks { get; }
}
