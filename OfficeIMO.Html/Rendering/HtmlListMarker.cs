namespace OfficeIMO.Html;

internal sealed class HtmlListMarker {
    internal HtmlListMarker(string content, HtmlRenderBoxStyle style, string position, HtmlRenderFlowBlock? image = null) {
        Content = content ?? throw new ArgumentNullException(nameof(content));
        Style = style ?? throw new ArgumentNullException(nameof(style));
        Position = position;
        Image = image;
    }

    internal string Content { get; }
    internal HtmlRenderBoxStyle Style { get; }
    internal string Position { get; }
    internal HtmlRenderFlowBlock? Image { get; }
    internal bool IsImage => Image != null;
    internal bool IsOutside => string.Equals(Position, "outside", StringComparison.OrdinalIgnoreCase);
}
