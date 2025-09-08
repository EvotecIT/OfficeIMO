namespace OfficeIMO.Markdown;

public sealed class H1 : IMarkdownBlock {
    private readonly HeadingBlock _h;
    public H1(string text) { _h = new HeadingBlock(1, text); }
    public string RenderMarkdown() => _h.RenderMarkdown();
    public string RenderHtml() => _h.RenderHtml();
}

