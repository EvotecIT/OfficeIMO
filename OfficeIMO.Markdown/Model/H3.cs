namespace OfficeIMO.Markdown;

public sealed class H3 : IMarkdownBlock {
    private readonly HeadingBlock _h;
    public H3(string text) { _h = new HeadingBlock(3, text); }
    public string RenderMarkdown() => _h.RenderMarkdown();
    public string RenderHtml() => _h.RenderHtml();
}

