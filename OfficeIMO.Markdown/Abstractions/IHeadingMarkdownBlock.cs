namespace OfficeIMO.Markdown;

internal interface IHeadingMarkdownBlock {
    int Level { get; }
    string Text { get; }
}
