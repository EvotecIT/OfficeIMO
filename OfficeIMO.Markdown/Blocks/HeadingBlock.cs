using System;

namespace OfficeIMO.Markdown;

public sealed class HeadingBlock : IMarkdownBlock {
    public int Level { get; }
    public string Text { get; }
    public HeadingBlock(int level, string text) {
        Level = Math.Clamp(level, 1, 6);
        Text = text ?? string.Empty;
    }
    public string RenderMarkdown() => new string('#', Level) + " " + Text;
    public string RenderHtml() => $"<h{Level}>{System.Net.WebUtility.HtmlEncode(Text)}</h{Level}>";
}

