namespace OfficeIMO.Markdown;

public sealed class H2 : IMarkdownBlock {
    private readonly HeadingBlock _h;
    public H2(string text) { _h = new HeadingBlock(2, text); }
    public string RenderMarkdown() => _h.RenderMarkdown();
    public string RenderHtml() => _h.RenderHtml();
}

