using AngleSharp.Dom;
using AngleSharp.Dom.Events;
using AngleSharp.Html.Dom;
using AngleSharp.Html.Dom.Events;

namespace OfficeIMO.Html.Runtime.Worker;

internal sealed partial class RuntimeAutomation {
    private readonly IDocument _document;
    private readonly RuntimeFocusController _focus;
    private readonly RuntimeHistoryBindings _history;
    private readonly RuntimeLocatorResolver _locators;
    private readonly RuntimeViewport _viewport;

    internal RuntimeAutomation(IDocument document, HtmlScriptRequest options, RuntimeFocusController focus,
        RuntimeHistoryBindings history, RuntimeViewport viewport, Jint.Engine engine) {
        _document = document;
        _engine = engine;
        _options = options;
        _focus = focus;
        _history = history;
        _locators = new RuntimeLocatorResolver(document, options);
        _viewport = viewport;
    }

    internal RuntimeViewport Viewport => _viewport;

    internal HtmlAutomationResult Run(HtmlAutomationRequest request, CancellationToken token) {
        if (!_viewport.Enabled && (request.Action is HtmlAutomationAction.Hover or HtmlAutomationAction.Press or HtmlAutomationAction.ScrollIntoView
            || request.Action == HtmlAutomationAction.Wait && request.WaitState is HtmlLocatorWaitState.Visible or HtmlLocatorWaitState.Hidden or HtmlLocatorWaitState.InViewport))
            return Failure(HtmlAutomationStatus.Unsupported, "Layout and input operations require WebApplicationV1.");
        IReadOnlyList<IElement> matches;
        try { matches = _locators.Resolve(request.Query, token); }
        catch (DomException error) { return Failure(HtmlAutomationStatus.InvalidLocator, error.Message); }
        if (request.Action == HtmlAutomationAction.Count) return new() { MatchCount = matches.Count };
        if (request.Action == HtmlAutomationAction.Wait
            && request.WaitState is (HtmlLocatorWaitState.Detached or HtmlLocatorWaitState.Hidden)
            && matches.Count == 0)
            return new();
        if (request.Action == HtmlAutomationAction.Wait && request.WaitState == HtmlLocatorWaitState.Detached)
            return matches.Count == 0 ? new() : Failure(HtmlAutomationStatus.NotReady, "The locator still matches attached elements.", matches.Count);
        if (matches.Count == 0) return Failure(HtmlAutomationStatus.NotFound, "The locator matches no attached element.");
        if (matches.Count != 1) return Failure(HtmlAutomationStatus.Ambiguous, "The operation requires exactly one element. Scope the query or select an explicit index.", matches.Count);
        IElement element = matches[0];
        RuntimeElementLayout layout = _viewport.Measure(element, token);
        HtmlRuntimeElementState inspected = Inspect(element, layout);
        if (request.Action == HtmlAutomationAction.Inspect) return Success(inspected);
        if (request.Action == HtmlAutomationAction.Wait) {
            bool ready = request.WaitState switch {
                HtmlLocatorWaitState.Attached => true,
                HtmlLocatorWaitState.Enabled => !inspected.IsDisabled,
                HtmlLocatorWaitState.Disabled => inspected.IsDisabled,
                HtmlLocatorWaitState.Editable => inspected.IsEditable,
                HtmlLocatorWaitState.Focused => inspected.IsFocused,
                HtmlLocatorWaitState.Visible => inspected.IsVisible == true,
                HtmlLocatorWaitState.Hidden => inspected.IsVisible == false,
                HtmlLocatorWaitState.InViewport => inspected.IsInViewport == true,
                HtmlLocatorWaitState.Value => inspected.Value != null && inspected.Value == request.Value,
                HtmlLocatorWaitState.Text => inspected.Text == RuntimeLocatorResolver.Normalize(request.Value),
                HtmlLocatorWaitState.Checked => inspected.IsChecked != null && inspected.IsChecked == request.Checked,
                _ => false
            };
            return ready ? Success(inspected) : Failure(HtmlAutomationStatus.NotReady, "The requested element state has not been reached.", 1, inspected);
        }
        if (request.Action == HtmlAutomationAction.Blur) { _focus.Blur(element); return Success(element); }
        if (request.Action == HtmlAutomationAction.ScrollIntoView) {
            if (_viewport.Enabled && !layout.IsVisible)
                return Failure(HtmlAutomationStatus.NotReady, "The element has no visible layout box.", 1, inspected);
            return Success(element, _viewport.ScrollIntoView(element, token));
        }
        if (RuntimeFocusController.Disabled(element) || RuntimeFocusController.HiddenByMarkup(element))
            return Failure(HtmlAutomationStatus.NotReady, "The element is disabled, hidden by markup or inside an inert subtree.", 1, inspected);
        if (_viewport.Enabled) {
            if (!layout.IsVisible) return Failure(HtmlAutomationStatus.NotReady, "The element has no visible layout box.", 1, inspected);
            if (request.Action is (HtmlAutomationAction.Click or HtmlAutomationAction.SetChecked or HtmlAutomationAction.Hover) && !layout.AcceptsPointerEvents)
                return Failure(HtmlAutomationStatus.NotReady, "The element does not accept pointer events.", 1, inspected);
            if (!layout.IsInViewport) layout = _viewport.ScrollIntoView(element, token);
        }
        return request.Action switch {
            HtmlAutomationAction.Focus => RuntimeFocusController.CanFocus(element)
                ? _focus.Focus(element) ? Success(element) : Failure(HtmlAutomationStatus.Rejected, "Page handlers redirected focus.", 1, Inspect(element))
                : Failure(HtmlAutomationStatus.Unsupported, "This element is not focusable.", 1, Inspect(element)),
            HtmlAutomationAction.Fill => Fill(element, request.Value!),
            HtmlAutomationAction.SelectOptions => Select(element, request.Values),
            HtmlAutomationAction.SetChecked => SetChecked(element, request.Checked!.Value, layout),
            HtmlAutomationAction.Click => _viewport.Enabled ? PointerClick(element, layout) : Activate(element, focusTarget: true),
            HtmlAutomationAction.Hover => _viewport.Enabled ? Hover(element, layout) : Failure(HtmlAutomationStatus.Unsupported, "Hover requires WebApplicationV1.", 1, inspected),
            HtmlAutomationAction.Press => _viewport.Enabled ? Press(element, request.Value!) : Failure(HtmlAutomationStatus.Unsupported, "Press requires WebApplicationV1.", 1, inspected),
            _ => Failure(HtmlAutomationStatus.Unsupported, "Unsupported automation operation.", 1)
        };
    }

    internal HtmlRuntimeElementState Inspect(IElement element, CancellationToken token = default) =>
        Inspect(element, _viewport.Measure(element, token));

    private HtmlRuntimeElementState Inspect(IElement element, RuntimeElementLayout layout) => new() {
        ElementName = element.LocalName, Id = element.Id ?? string.Empty,
        AccessibleName = RuntimeLocatorResolver.AccessibleName(element),
        Text = RuntimeLocatorResolver.Normalize(element.TextContent), Value = RuntimeFocusController.Value(element),
        SelectedValues = element is IHtmlSelectElement select ? Array.AsReadOnly(select.Options.Where(option => option.IsSelected).Select(option => option.Value).ToArray()) : Array.Empty<string>(),
        IsChecked = element is IHtmlInputElement check && check.Type is "checkbox" or "radio" ? check.IsChecked : null,
        IsIndeterminate = element is IHtmlInputElement mixed && mixed.Type == "checkbox" && mixed.IsIndeterminate,
        IsDisabled = RuntimeFocusController.Disabled(element), IsReadOnly = ReadOnly(element), IsEditable = SupportsFill(element) && ReadyToEdit(element),
        IsHiddenByMarkup = RuntimeFocusController.HiddenByMarkup(element), IsFocused = ReferenceEquals(_focus.Focused, element), IsConnected = RuntimeFocusController.IsConnected(element),
        IsVisible = _viewport.Enabled ? layout.IsVisible : null,
        IsInViewport = _viewport.Enabled && layout.IsVisible ? layout.IsInViewport : null,
        AcceptsPointerEvents = _viewport.Enabled ? layout.AcceptsPointerEvents : null,
        BoundingBox = layout.Box,
        ScrollX = layout.ScrollX,
        ScrollY = layout.ScrollY
    };

    private HtmlAutomationResult Fill(IElement element, string value) {
        if (!SupportsFill(element)) return Failure(HtmlAutomationStatus.Unsupported, "Fill supports text, search, email, URL, telephone and password inputs and textareas.", 1, Inspect(element));
        if (!ReadyToEdit(element)) return Failure(HtmlAutomationStatus.NotReady, "The text control is not editable.", 1, Inspect(element));
        if (!_focus.Focus(element) || !SupportsFill(element) || !ReadyToEdit(element)) return Failure(HtmlAutomationStatus.Rejected, "Focus handlers changed the editing target.", 1, Inspect(element));
        var beforeInput = new InputEvent("beforeinput", true, true, value);
        element.Dispatch(beforeInput);
        if (beforeInput.IsDefaultPrevented) return Failure(HtmlAutomationStatus.Rejected, "The page cancelled beforeinput.", 1, Inspect(element));
        if (!SupportsFill(element) || !ReadyToEdit(element) || !ReferenceEquals(_focus.Focused, element)) return Failure(HtmlAutomationStatus.Rejected, "Beforeinput handlers changed the editing target.", 1, Inspect(element));
        _focus.ChangedByUser(element);
        if (element is IHtmlInputElement input) {
            string normalized = value.Replace("\r", string.Empty).Replace("\n", string.Empty);
            input.Value = input.Type is "email" or "url" ? normalized.Trim(' ', '\t', '\f') : normalized;
        }
        else ((IHtmlTextAreaElement)element).Value = value;
        element.Dispatch(new InputEvent("input", true, false, value));
        return Success(element);
    }

    private HtmlAutomationResult Select(IElement element, IReadOnlyList<string> values) {
        if (element is not IHtmlSelectElement select) return Failure(HtmlAutomationStatus.Unsupported, "SelectOptions requires a select element.", 1, Inspect(element));
        string[] requested = values.Distinct(StringComparer.Ordinal).ToArray();
        if (!select.IsMultiple && requested.Length > 1) return Failure(HtmlAutomationStatus.InvalidValue, "A dropdown accepts at most one selected option.", 1);
        var selected = new HashSet<IHtmlOptionElement>();
        foreach (string value in requested) {
            IHtmlOptionElement[] candidates = select.Options.Where(option => string.Equals(option.Value, value, StringComparison.Ordinal)).ToArray();
            if (candidates.Length != 1) return Failure(HtmlAutomationStatus.InvalidValue, "Each requested option value must identify exactly one option.", 1);
            if (HtmlFormControlSemantics.IsOptionEffectivelyDisabled(candidates[0])) return Failure(HtmlAutomationStatus.NotReady, "A requested option is disabled.", 1);
            selected.Add(candidates[0]);
        }
        if (!_focus.Focus(element) || !RuntimeFocusController.IsConnected(element) || RuntimeFocusController.Disabled(element) || RuntimeFocusController.HiddenByMarkup(element))
            return Failure(HtmlAutomationStatus.Rejected, "Focus handlers changed the selection target.", 1, Inspect(element));
        // A focus handler may replace or disable an option. Validate the retained references before mutation.
        if (!select.IsMultiple && requested.Length > 1 || selected.Any(option => !select.Options.Contains(option)
            || HtmlFormControlSemantics.IsOptionEffectivelyDisabled(option))
            || requested.Any(value => {
                var current = select.Options.Where(option => string.Equals(option.Value, value, StringComparison.Ordinal)).ToArray();
                return current.Length != 1 || !selected.Contains(current[0]);
            }))
            return Failure(HtmlAutomationStatus.Rejected, "Focus handlers changed the requested options.", 1, Inspect(element));
        bool changed = select.Options.Any(option => option.IsSelected != selected.Contains(option));
        foreach (IHtmlOptionElement option in select.Options) option.IsSelected = selected.Contains(option);
        if (changed) { element.Dispatch(new Event("input", true, false)); element.Dispatch(new Event("change", true, false)); }
        return Success(element);
    }

    private HtmlAutomationResult SetChecked(IElement element, bool value, RuntimeElementLayout layout) {
        if (element is not IHtmlInputElement input || input.Type is not ("checkbox" or "radio"))
            return Failure(HtmlAutomationStatus.Unsupported, "SetChecked requires a checkbox or radio input.", 1, Inspect(element));
        if (input.Type == "radio" && !value) return Failure(HtmlAutomationStatus.InvalidValue, "A radio is unchecked by choosing another radio in its group.", 1);
        if (input.IsChecked == value) return Success(element);
        HtmlAutomationResult result = _viewport.Enabled ? PointerClick(element, layout) : Activate(element, true);
        if (result.Status != HtmlAutomationStatus.Success) return result;
        return input.IsChecked == value ? Success(element)
            : Failure(HtmlAutomationStatus.Rejected, "The page prevented the requested checked state.", 1, Inspect(element));
    }

    internal HtmlAutomationResult Activate(IElement element, bool focusTarget, RuntimeElementLayout? pointerLayout = null) {
        if (element is not IHtmlElement) return Failure(HtmlAutomationStatus.Unsupported, "DOM activation requires an HTML element.", 1);
        if (RuntimeFocusController.Disabled(element)) return Failure(HtmlAutomationStatus.NotReady, "The element is disabled.", 1, Inspect(element));
        if (focusTarget && RuntimeFocusController.CanFocus(element) && !_focus.Focus(element))
            return Failure(HtmlAutomationStatus.Rejected, "Page handlers redirected focus.", 1, Inspect(element));
        if (RuntimeFocusController.Disabled(element) || focusTarget && (!RuntimeFocusController.IsConnected(element) || RuntimeFocusController.HiddenByMarkup(element))) return Failure(HtmlAutomationStatus.Rejected, "The activation target changed during focus.", 1, Inspect(element));
        var anchor = element is IHtmlButtonElement or IHtmlInputElement ? null : element.Closest("a[href]") as IHtmlAnchorElement;
        var check = element as IHtmlInputElement;
        bool isChoice = check?.Type is "checkbox" or "radio";
        var saved = new Dictionary<IHtmlInputElement, bool>();
        bool mixed = check?.IsIndeterminate == true;
        if (isChoice) {
            saved.Add(check!, check!.IsChecked);
            if (check.Type == "radio" && !string.IsNullOrEmpty(check.Name)) {
                var owner = HtmlFormControlSemantics.ResolveFormOwner(check);
                INode root=check;
                while(root.Parent!=null)root=root.Parent;
                var peers=root is IParentNode parent ? parent.QuerySelectorAll("input").OfType<IHtmlInputElement>() : Enumerable.Empty<IHtmlInputElement>();
                if(root is IHtmlInputElement rootInput)peers=peers.Prepend(rootInput);
                foreach (var other in peers)
                    if (!ReferenceEquals(other, check) && other.Type == "radio" && other.Name == check.Name && ReferenceEquals(HtmlFormControlSemantics.ResolveFormOwner(other), owner))
                        saved.Add(other, other.IsChecked);
                foreach (var other in saved.Keys) other.IsChecked = ReferenceEquals(other, check);
            } else check.IsChecked = check.Type == "radio" || !check.IsChecked;
            check.IsIndeterminate = false;
        }
        int x = (int)Math.Round((pointerLayout?.Box?.X ?? 0D) + (pointerLayout?.Box?.Width ?? 0D) / 2D);
        int y = (int)Math.Round((pointerLayout?.Box?.Y ?? 0D) + (pointerLayout?.Box?.Height ?? 0D) / 2D);
        var click = new MouseEvent();
        click.Init("click", true, true, _document.DefaultView, pointerLayout == null && !focusTarget ? 0 : 1,
            x, y, x, y, false, false, false, false, MouseButton.Primary, null);
        element.Dispatch(click);
        if (click.IsDefaultPrevented) {
            foreach (var pair in saved) pair.Key.IsChecked = pair.Value;
            if (isChoice) check!.IsIndeterminate = mixed;
            return Success(element);
        }
        if (isChoice && saved[check!] != check!.IsChecked) {
            element.Dispatch(new Event("input", true, false)); element.Dispatch(new Event("change", true, false));
        }
        string type = HtmlFormControlSemantics.GetEffectiveType(element.LocalName, element.GetAttribute("type"));
        if (anchor != null && anchor.HasAttribute("href")) {
            string target = anchor.GetAttribute("target") ?? _document.QuerySelector("base[target]")?.GetAttribute("target") ?? "";
            if (anchor.HasAttribute("download") || target.Length > 0 && target.ToLowerInvariant() is not ("_self" or "_top" or "_parent"))
                return Failure(HtmlAutomationStatus.Unsupported,"Downloads and additional browsing contexts are outside this interaction profile.",1,Inspect(element));
            try { RuntimeDocumentUrls.Base(_document); _history.NavigateFragment(anchor.Href); }
            catch (Jint.Runtime.JavaScriptException error) { return Failure(HtmlAutomationStatus.Unsupported,error.Message,1,Inspect(element)); }
        }
        if (element is IHtmlButtonElement or IHtmlInputElement && type is "submit" or "reset" && HtmlFormControlSemantics.ResolveFormOwner(element) is IHtmlFormElement form)
            return ApplyFormDefault(form, (IHtmlElement)element, type);
        return Success(element);
    }

    private static bool SupportsFill(IElement element) => element is IHtmlTextAreaElement
        || element is IHtmlInputElement input && input.Type is "text" or "search" or "email" or "url" or "tel" or "password";
    private static bool ReadOnly(IElement element) => element.HasAttribute("readonly")
        && HtmlFormControlSemantics.IsReadOnlyStateApplicable(element.LocalName, HtmlFormControlSemantics.GetEffectiveType(element.LocalName, element.GetAttribute("type")));
    private static bool ReadyToEdit(IElement element) => RuntimeFocusController.IsConnected(element) && !ReadOnly(element) && !RuntimeFocusController.Disabled(element) && !RuntimeFocusController.HiddenByMarkup(element);
    private HtmlAutomationResult Success(IElement element) => new() { MatchCount = 1, Element = Inspect(element) };
    private HtmlAutomationResult Success(IElement element, RuntimeElementLayout layout) => new() { MatchCount = 1, Element = Inspect(element, layout) };
    private static HtmlAutomationResult Success(HtmlRuntimeElementState state) => new() { MatchCount = 1, Element = state };
    private static HtmlAutomationResult Failure(HtmlAutomationStatus status, string message, int count = 0, HtmlRuntimeElementState? state = null) => new() { Status = status, Message = message, MatchCount = count, Element = state };
}
