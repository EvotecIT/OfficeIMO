namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    // Anonymous blocks after a preceding block are not the parent's first
    // formatted line. Descendant block elements still inherit independently.
    private static HtmlRenderBoxStyle WithoutTextIndent(HtmlRenderBoxStyle style) {
        if (style.TextIndent == null) return style;
        HtmlRenderBoxStyle result = style.Clone();
        result.TextIndent = null;
        return result;
    }
}
