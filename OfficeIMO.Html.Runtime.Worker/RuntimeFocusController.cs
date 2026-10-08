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
    private RuntimeViewport? _viewport;
    private bool _measuringVisibility;
    internal IElement? Current => _focused;
    internal void UseViewport(RuntimeViewport viewport) => _viewport = viewport;
    internal void Reset() {
        _focused = null;
        _initialValue = null;
        _dirty = false;
        _viewport = null;
        _transition++;
    }

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
        if (_transition != transition) return false;
        _focused = target;
        // CSS eligibility must use the destination focus state. A transient
        // gap between blur and focus would hide :focus-within controls even
        // when both the old and new focused elements belong to that subtree.
        if (target != null && !CanFocus(target)) {
            _focused = null;
            return false;
        }
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
    internal static string? ObservableValue(IElement element) =>
        element is IHtmlInputElement input && string.Equals(input.Type, "password", StringComparison.OrdinalIgnoreCase)
            ? null
            : Value(element);
    internal static bool ExposesTextSelection(IElement element) =>
        element is not IHtmlInputElement input || !string.Equals(input.Type, "password", StringComparison.OrdinalIgnoreCase);
    internal static bool HiddenByMarkup(IElement element) {
        for (IElement? current = element; current != null; current = current.ParentElement)
            if (current.HasAttribute("hidden") || current.HasAttribute("inert")) return true;
        return false;
    }
    internal static bool Disabled(IElement element) => element is IHtmlInputElement or IHtmlTextAreaElement or IHtmlSelectElement or IHtmlButtonElement
        && HtmlFormControlSemantics.IsEffectivelyDisabled(element);
    internal bool CanFocus(IElement element) => element is IHtmlElement && IsConnected(element) && !Disabled(element) && !HiddenByMarkup(element)
        && (element.HasAttribute("tabindex") || element is IHtmlTextAreaElement or IHtmlSelectElement or IHtmlButtonElement
            || element is IHtmlInputElement input && input.Type != "hidden"
            || element is IHtmlAnchorElement && element.HasAttribute("href")) && HasVisibleLayout(element);

    private bool HasVisibleLayout(IElement element) {
        if (_viewport?.Enabled != true || _measuringVisibility) return true;
        // Computing a layout can match :focus selectors, which read this owner's
        // current focus. Keep that nested read from starting another layout.
        _measuringVisibility = true;
        try { return _viewport.Measure(element, _viewport.CurrentCommandToken, measureGeometry: false).IsCssVisible; }
        finally { _measuringVisibility = false; }
    }

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
