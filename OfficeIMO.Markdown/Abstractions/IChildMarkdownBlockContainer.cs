namespace OfficeIMO.Markdown;

internal interface IChildMarkdownBlockContainer {
    IReadOnlyList<IMarkdownBlock> ChildBlocks { get; }
}
