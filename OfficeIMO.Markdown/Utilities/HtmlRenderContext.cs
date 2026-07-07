namespace OfficeIMO.Markdown;

/// <summary>
/// Ambient render context so nested blocks (quotes/details/lists) can obey HTML rendering policies.
/// </summary>
internal static class HtmlRenderContext {
    private static readonly System.Threading.AsyncLocal<HtmlOptions?> s_Options = new();
    private static readonly System.Threading.AsyncLocal<HtmlFootnoteRenderState?> s_Footnotes = new();
    private static readonly System.Threading.AsyncLocal<MarkdownBodyRenderContext?> s_BodyContext = new();
    private static readonly System.Threading.AsyncLocal<int> s_RenderListItemAttributesDepth = new();

    internal static HtmlOptions? Options => s_Options.Value;
    internal static HtmlFootnoteRenderState? Footnotes => s_Footnotes.Value;
    internal static MarkdownBodyRenderContext? BodyContext => s_BodyContext.Value;
    internal static bool RenderListItemAttributes => s_RenderListItemAttributesDepth.Value > 0;

    internal static IDisposable Push(HtmlOptions options, HtmlFootnoteRenderState? footnotes = null) {
        var prior = s_Options.Value;
        var priorFootnotes = s_Footnotes.Value;
        s_Options.Value = options;
        s_Footnotes.Value = footnotes;
        return new Popper(prior, priorFootnotes);
    }

    internal static IDisposable PushBodyContext(MarkdownBodyRenderContext context) {
        var prior = s_BodyContext.Value;
        s_BodyContext.Value = context;
        return new BodyContextPopper(prior);
    }

    internal static IDisposable PushRenderListItemAttributes() {
        s_RenderListItemAttributesDepth.Value++;
        return new RenderListItemAttributesPopper();
    }

    private readonly struct Popper : IDisposable {
        private readonly HtmlOptions? _prior;
        private readonly HtmlFootnoteRenderState? _priorFootnotes;

        public Popper(HtmlOptions? prior, HtmlFootnoteRenderState? priorFootnotes) {
            _prior = prior;
            _priorFootnotes = priorFootnotes;
        }

        public void Dispose() {
            s_Options.Value = _prior;
            s_Footnotes.Value = _priorFootnotes;
        }
    }

    private readonly struct BodyContextPopper : IDisposable {
        private readonly MarkdownBodyRenderContext? _prior;

        public BodyContextPopper(MarkdownBodyRenderContext? prior) {
            _prior = prior;
        }

        public void Dispose() {
            s_BodyContext.Value = _prior;
        }
    }

    private readonly struct RenderListItemAttributesPopper : IDisposable {
        public void Dispose() {
            var current = s_RenderListItemAttributesDepth.Value;
            s_RenderListItemAttributesDepth.Value = current > 0 ? current - 1 : 0;
        }
    }
}
