namespace OfficeIMO.Markdown;

/// <summary>
/// Ambient render context so nested blocks (quotes/details/lists) can obey HTML rendering policies.
/// </summary>
/// <remarks>
/// HTML rendering runs synchronously on one thread and every push is undone before the render returns,
/// so the context lives in thread-static fields; changing an async-local value allocates on every push and pop.
/// </remarks>
internal static class HtmlRenderContext {
    [ThreadStatic] private static MarkdownHtmlExportOptions? t_Options;
    [ThreadStatic] private static HtmlFootnoteRenderState? t_Footnotes;
    [ThreadStatic] private static MarkdownBodyRenderContext? t_BodyContext;
    [ThreadStatic] private static int t_RenderListItemAttributesDepth;

    internal static MarkdownHtmlExportOptions? Options => t_Options;
    internal static HtmlFootnoteRenderState? Footnotes => t_Footnotes;
    internal static MarkdownBodyRenderContext? BodyContext => t_BodyContext;
    internal static bool RenderListItemAttributes => t_RenderListItemAttributesDepth > 0;

    internal static Popper Push(MarkdownHtmlExportOptions options, HtmlFootnoteRenderState? footnotes = null) {
        var prior = t_Options;
        var priorFootnotes = t_Footnotes;
        t_Options = options;
        t_Footnotes = footnotes;
        return new Popper(prior, priorFootnotes);
    }

    internal static BodyContextPopper PushBodyContext(MarkdownBodyRenderContext context) {
        var prior = t_BodyContext;
        t_BodyContext = context;
        return new BodyContextPopper(prior);
    }

    internal static RenderListItemAttributesPopper PushRenderListItemAttributes() {
        t_RenderListItemAttributesDepth++;
        return new RenderListItemAttributesPopper();
    }

    internal readonly struct Popper : IDisposable {
        private readonly MarkdownHtmlExportOptions? _prior;
        private readonly HtmlFootnoteRenderState? _priorFootnotes;

        public Popper(MarkdownHtmlExportOptions? prior, HtmlFootnoteRenderState? priorFootnotes) {
            _prior = prior;
            _priorFootnotes = priorFootnotes;
        }

        public void Dispose() {
            t_Options = _prior;
            t_Footnotes = _priorFootnotes;
        }
    }

    internal readonly struct BodyContextPopper : IDisposable {
        private readonly MarkdownBodyRenderContext? _prior;

        public BodyContextPopper(MarkdownBodyRenderContext? prior) {
            _prior = prior;
        }

        public void Dispose() {
            t_BodyContext = _prior;
        }
    }

    internal readonly struct RenderListItemAttributesPopper : IDisposable {
        public void Dispose() {
            var current = t_RenderListItemAttributesDepth;
            t_RenderListItemAttributesDepth = current > 0 ? current - 1 : 0;
        }
    }
}
