using AngleSharp;
using AngleSharp.Css;
using AngleSharp.Css.Dom;
using AngleSharp.Dom;

namespace OfficeIMO.Html;

// A measurement clone has its own selector context. Transfer focus by source
// identity without changing the live document or the shared default factories.
internal sealed class HtmlInteractionFocusState : IDisposable {
    internal IElement? Focused { get; set; }
    internal IBrowsingContext Context { get; }

    internal HtmlInteractionFocusState() {
        var selectors = new DefaultPseudoClassSelectorFactory();
        selectors.Unregister("focus");
        selectors.Unregister("focus-within");
        selectors.Register("focus", new FocusSelector(this, false));
        selectors.Register("focus-within", new FocusSelector(this, true));
        Context = BrowsingContext.New(Configuration.Default.WithOnly<IPseudoClassSelectorFactory>(selectors));
    }

    public void Dispose() => Context.Dispose();

    private sealed class FocusSelector(HtmlInteractionFocusState owner, bool within) : ISelector {
        public string Text => within ? ":focus-within" : ":focus";
        public Priority Specificity => Priority.OneClass;
        public bool Match(IElement element, IElement? scope) {
            for (IElement? current = owner.Focused; current != null; current = within ? current.ParentElement : null)
                if (ReferenceEquals(element, current)) return true;
            return false;
        }
        public void Accept(ISelectorVisitor visitor) => visitor.PseudoClass(within ? "focus-within" : "focus");
    }
}
