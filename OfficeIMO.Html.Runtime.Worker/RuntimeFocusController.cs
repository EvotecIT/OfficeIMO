using AngleSharp.Css;
using AngleSharp.Css.Dom;
using AngleSharp.Dom;
using AngleSharp.Dom.Events;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html.Runtime.Worker;

// The retained DOM exposes focus methods as stubs. Keep focus, selector matching and JS getters
// in one session owner; no private provider fields or reflection are needed.
internal sealed class RuntimeFocusController {
    private IElement? _focused;
    private string? _initialValue;
    private bool _dirty;
    private long _transition;
    internal IElement? Focused {
        get {
            if (_focused != null && !CanFocus(_focused)) {
                _focused = null;
                _initialValue = null;
                _dirty = false;
            }
            return _focused;
        }
    }

    internal static bool IsConnected(INode element) {
        INode root = element;
        while (root.Parent != null) root = root.Parent;
        return root is IDocument;
    }

    internal DefaultPseudoClassSelectorFactory CreateSelectors() {
        var selectors = new DefaultPseudoClassSelectorFactory();
        selectors.Unregister("focus");
        selectors.Unregister("focus-within");
        selectors.Register("focus", new FocusSelector(this, false));
        selectors.Register("focus-within", new FocusSelector(this, true));
        return selectors;
    }

    internal bool Focus(IElement? target) {
        if (target != null && !CanFocus(target)) return false;
        IElement? old = Focused;
        if (ReferenceEquals(old, target)) return true;
        long transition = ++_transition;
        _focused = null;
        bool changed = _dirty && old != null && !string.Equals(_initialValue, Value(old), StringComparison.Ordinal);
        _dirty = false; _initialValue = null;
        if (old != null) {
            if (changed) old.Dispatch(new Event("change", true, false));
            if (_transition != transition) return false;
            old.Dispatch(new FocusEvent("blur", false, false, old.Owner!.DefaultView, 0, target));
            if (_transition != transition) return false;
            old.Dispatch(new FocusEvent("focusout", true, false, old.Owner!.DefaultView, 0, target));
        }
        if (_transition != transition || target != null && !CanFocus(target)) return false;
        _focused = target;
        _initialValue = target == null ? null : Value(target);
        if (target != null) {
            target.Dispatch(new FocusEvent("focus", false, false, target.Owner!.DefaultView, 0, old));
            if (_transition != transition || !ReferenceEquals(Focused, target)) return false;
            target.Dispatch(new FocusEvent("focusin", true, false, target.Owner!.DefaultView, 0, old));
        }
        return ReferenceEquals(Focused, target);
    }

    internal void Blur(IElement element) { if (ReferenceEquals(Focused, element)) Focus(null); }
    internal void ChangedByUser(IElement element) { if (ReferenceEquals(Focused, element)) _dirty = true; }
    internal static string? Value(IElement element) => element switch {
        IHtmlInputElement input => input.Value,
        IHtmlTextAreaElement area => area.Value,
        IHtmlSelectElement select => select.Value ?? string.Empty,
        _ => null
    };
    internal static bool HiddenByMarkup(IElement element) {
        for (IElement? current = element; current != null; current = current.ParentElement)
            if (current.HasAttribute("hidden") || current.HasAttribute("inert")) return true;
        return false;
    }
    internal static bool Disabled(IElement element) => element is IHtmlInputElement or IHtmlTextAreaElement or IHtmlSelectElement or IHtmlButtonElement
        && HtmlFormControlSemantics.IsEffectivelyDisabled(element);
    internal static bool CanFocus(IElement element) => element is IHtmlElement && IsConnected(element) && !Disabled(element) && !HiddenByMarkup(element)
        && (element.HasAttribute("tabindex") || element is IHtmlTextAreaElement or IHtmlSelectElement or IHtmlButtonElement
            || element is IHtmlInputElement input && input.Type != "hidden"
            || element is IHtmlAnchorElement && element.HasAttribute("href"));

    private sealed class FocusSelector(RuntimeFocusController owner, bool within) : ISelector {
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
