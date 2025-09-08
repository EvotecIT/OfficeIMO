namespace OfficeIMO.Markdown;

public sealed class UnorderedList : IMarkdownBlock {
    private readonly UnorderedListBlock _ul = new UnorderedListBlock();
    public void Add(ListItem item) => _ul.Items.Add(item);
    public string RenderMarkdown() => _ul.RenderMarkdown();
    public string RenderHtml() => _ul.RenderHtml();
}

