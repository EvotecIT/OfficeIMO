namespace OfficeIMO.Markdown;

/// <summary>
/// Marks an inline node as a container of nested inline content.
/// Implement this on custom container nodes so source-span binding, syntax tree generation, and
/// object-tree traversal can descend into the nested inline sequence.
/// </summary>
public interface IInlineContainerMarkdownInline {
    /// <summary>
    /// Nested inline content owned by the container node.
    /// </summary>
    InlineSequence? NestedInlines { get; }
}
