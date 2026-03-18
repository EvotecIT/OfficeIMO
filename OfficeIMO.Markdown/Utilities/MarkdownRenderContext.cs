namespace OfficeIMO.Markdown;

/// <summary>
/// Ambient render context so nested markdown renderers can honor active markdown write options.
/// </summary>
internal static class MarkdownRenderContext {
    private static readonly System.Threading.AsyncLocal<MarkdownWriteOptions?> s_Options = new();

    internal static MarkdownWriteOptions? Options => s_Options.Value;

    internal static IDisposable Push(MarkdownWriteOptions options) {
        var prior = s_Options.Value;
        s_Options.Value = options;
        return new Popper(prior);
    }

    private readonly struct Popper : System.IDisposable {
        private readonly MarkdownWriteOptions? _prior;

        public Popper(MarkdownWriteOptions? prior) {
            _prior = prior;
        }

        public void Dispose() {
            s_Options.Value = _prior;
        }
    }
}
