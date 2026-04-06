namespace OfficeIMO.Markdown;

/// <summary>
/// Ambient render context so nested blocks (quotes/details/lists) can obey HTML rendering policies.
/// </summary>
internal static class HtmlRenderContext {
    private static readonly System.Threading.AsyncLocal<HtmlOptions?> s_Options = new();
    private static readonly System.Threading.AsyncLocal<HtmlFootnoteRenderState?> s_Footnotes = new();

    internal static HtmlOptions? Options => s_Options.Value;
    internal static HtmlFootnoteRenderState? Footnotes => s_Footnotes.Value;

    internal static IDisposable Push(HtmlOptions options, HtmlFootnoteRenderState? footnotes = null) {
        var prior = s_Options.Value;
        var priorFootnotes = s_Footnotes.Value;
        s_Options.Value = options;
        s_Footnotes.Value = footnotes;
        return new Popper(prior, priorFootnotes);
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
}

