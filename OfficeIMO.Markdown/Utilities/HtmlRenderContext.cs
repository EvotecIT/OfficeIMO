namespace OfficeIMO.Markdown;

/// <summary>
/// Ambient render context so nested blocks (quotes/details/lists) can obey HTML rendering policies.
/// </summary>
/// <remarks>
/// A render sets one async-local frame, so work an extension starts on another thread during the
/// render still sees its options and link policies. Nested pushes within the render change that
/// frame and restore it on dispose instead of setting the async-local again, which allocates on
/// every change.
/// </remarks>
internal static class HtmlRenderContext {
    private static readonly System.Threading.AsyncLocal<Frame?> s_Frame = new();

    internal static MarkdownHtmlExportOptions? Options => s_Frame.Value?.Options;
    internal static HtmlFootnoteRenderState? Footnotes => s_Frame.Value?.Footnotes;
    internal static MarkdownBodyRenderContext? BodyContext => s_Frame.Value?.BodyContext;
    internal static bool RenderListItemAttributes => (s_Frame.Value?.RenderListItemAttributesDepth ?? 0) > 0;

    internal static Popper Push(MarkdownHtmlExportOptions options, HtmlFootnoteRenderState? footnotes = null) {
        var prior = s_Frame.Value;
        s_Frame.Value = new Frame(options, footnotes);
        return new Popper(prior);
    }

    internal static BodyContextPopper PushBodyContext(MarkdownBodyRenderContext context) {
        var frame = s_Frame.Value;
        if (frame == null) {
            frame = new Frame(options: null, footnotes: null);
            s_Frame.Value = frame;
            frame.BodyContext = context;
            return new BodyContextPopper(frame, prior: null, ownsFrame: true);
        }

        var prior = frame.BodyContext;
        frame.BodyContext = context;
        return new BodyContextPopper(frame, prior, ownsFrame: false);
    }

    internal static RenderListItemAttributesPopper PushRenderListItemAttributes() {
        var frame = s_Frame.Value;
        bool ownsFrame = frame == null;
        if (frame == null) {
            frame = new Frame(options: null, footnotes: null);
            s_Frame.Value = frame;
        }

        frame.RenderListItemAttributesDepth++;
        return new RenderListItemAttributesPopper(frame, ownsFrame);
    }

    internal sealed class Frame {
        internal Frame(MarkdownHtmlExportOptions? options, HtmlFootnoteRenderState? footnotes) {
            Options = options;
            Footnotes = footnotes;
        }

        internal MarkdownHtmlExportOptions? Options { get; }
        internal HtmlFootnoteRenderState? Footnotes { get; }
        internal MarkdownBodyRenderContext? BodyContext { get; set; }
        internal int RenderListItemAttributesDepth { get; set; }
    }

    internal readonly struct Popper : IDisposable {
        private readonly Frame? _prior;

        internal Popper(Frame? prior) {
            _prior = prior;
        }

        public void Dispose() {
            s_Frame.Value = _prior;
        }
    }

    internal readonly struct BodyContextPopper : IDisposable {
        private readonly Frame _frame;
        private readonly MarkdownBodyRenderContext? _prior;
        private readonly bool _ownsFrame;

        internal BodyContextPopper(Frame frame, MarkdownBodyRenderContext? prior, bool ownsFrame) {
            _frame = frame;
            _prior = prior;
            _ownsFrame = ownsFrame;
        }

        public void Dispose() {
            _frame.BodyContext = _prior;
            if (_ownsFrame) {
                s_Frame.Value = null;
            }
        }
    }

    internal readonly struct RenderListItemAttributesPopper : IDisposable {
        private readonly Frame _frame;
        private readonly bool _ownsFrame;

        internal RenderListItemAttributesPopper(Frame frame, bool ownsFrame) {
            _frame = frame;
            _ownsFrame = ownsFrame;
        }

        public void Dispose() {
            if (_frame.RenderListItemAttributesDepth > 0) {
                _frame.RenderListItemAttributesDepth--;
            }

            if (_ownsFrame) {
                s_Frame.Value = null;
            }
        }
    }
}
