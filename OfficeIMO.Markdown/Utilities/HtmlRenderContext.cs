namespace OfficeIMO.Markdown;

/// <summary>
/// Ambient render context so nested blocks (quotes/details/lists) can obey HTML rendering policies.
/// </summary>
internal static class HtmlRenderContext {
    private static readonly System.Threading.AsyncLocal<HtmlOptions?> s_Options = new();

    internal static HtmlOptions? Options => s_Options.Value;

    internal static IDisposable Push(HtmlOptions options) {
        var prior = s_Options.Value;
        s_Options.Value = options;
        return new Popper(prior);
    }

    private readonly struct Popper : IDisposable {
        private readonly HtmlOptions? _prior;
        public Popper(HtmlOptions? prior) { _prior = prior; }
        public void Dispose() { s_Options.Value = _prior; }
    }
}

