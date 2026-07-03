namespace OfficeIMO.Markdown;

/// <summary>
/// Ambient render context so nested markdown renderers can honor active markdown write options.
/// </summary>
internal static class MarkdownRenderContext {
    private static readonly System.Threading.AsyncLocal<MarkdownWriteOptions?> s_Options = new();
    private static readonly System.Threading.AsyncLocal<MarkdownWriteContext?> s_WriteContext = new();

    internal static MarkdownWriteOptions? Options => s_Options.Value;
    internal static MarkdownWriteContext? WriteContext => s_WriteContext.Value;

    internal static IDisposable Push(MarkdownWriteOptions options) {
        var prior = s_Options.Value;
        var priorContext = s_WriteContext.Value;
        s_Options.Value = options;
        s_WriteContext.Value = null;
        return new Popper(prior, priorContext);
    }

    internal static IDisposable Push(MarkdownWriteContext context) {
        var prior = s_Options.Value;
        var priorContext = s_WriteContext.Value;
        s_Options.Value = context.Options;
        s_WriteContext.Value = context;
        return new Popper(prior, priorContext);
    }

    private readonly struct Popper : System.IDisposable {
        private readonly MarkdownWriteOptions? _prior;
        private readonly MarkdownWriteContext? _priorContext;

        public Popper(MarkdownWriteOptions? prior, MarkdownWriteContext? priorContext) {
            _prior = prior;
            _priorContext = priorContext;
        }

        public void Dispose() {
            s_Options.Value = _prior;
            s_WriteContext.Value = _priorContext;
        }
    }
}
