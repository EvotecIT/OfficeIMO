using System.Globalization;
using AngleSharp.Dom;
using AngleSharp.Dom.Events;
using AngleSharp.Html.Dom;
using AngleSharp.Html.Dom.Events;

namespace OfficeIMO.Html.Runtime.Worker;

internal sealed partial class RuntimeAutomation {
    private IElement? _hovered;

    private HtmlAutomationResult Hover(IElement element, RuntimeElementLayout layout) {
        if (!layout.IsVisible || layout.Box == null)
            return Failure(HtmlAutomationStatus.NotReady, "The element has no visible layout box.", 1, Inspect(element, layout));
        if (!TryRefreshPointerTarget(element, out layout, out HtmlAutomationResult? initialFailure)) return initialFailure!;
        IElement? previous = _hovered;
        if (previous != null && !RuntimeFocusController.IsConnected(previous)) previous = null;
        if (!ReferenceEquals(previous, element)) {
            if (previous != null) {
                DispatchMouse(previous, "pointerout", true, true, layout, element);
                DispatchMouse(previous, "pointerleave", false, false, layout, element);
                DispatchMouse(previous, "mouseout", true, true, layout, element);
                DispatchMouse(previous, "mouseleave", false, false, layout, element);
            }
            if (!TryRefreshPointerTarget(element, out layout, out HtmlAutomationResult? arrivalFailure)) return arrivalFailure!;
            DispatchMouse(element, "pointerover", true, true, layout, previous);
            if (!TryRefreshPointerTarget(element, out layout, out arrivalFailure)) return arrivalFailure!;
            DispatchMouse(element, "pointerenter", false, false, layout, previous);
            if (!TryRefreshPointerTarget(element, out layout, out arrivalFailure)) return arrivalFailure!;
            DispatchMouse(element, "mouseover", true, true, layout, previous);
            if (!TryRefreshPointerTarget(element, out layout, out arrivalFailure)) return arrivalFailure!;
            DispatchMouse(element, "mouseenter", false, false, layout, previous);
            if (!TryRefreshPointerTarget(element, out layout, out arrivalFailure)) return arrivalFailure!;
            _hovered = element;
        }
        DispatchMouse(element, "pointermove", true, true, layout, null);
        if (!TryRefreshPointerTarget(element, out layout, out HtmlAutomationResult? moveFailure)) return moveFailure!;
        DispatchMouse(element, "mousemove", true, true, layout, null);
        if (!TryRefreshPointerTarget(element, out layout, out moveFailure)) return moveFailure!;
        return Success(element, layout);
    }

    private HtmlAutomationResult PointerClick(IElement element, RuntimeElementLayout layout) {
        HtmlAutomationResult hovered = Hover(element, layout);
        if (hovered.Status != HtmlAutomationStatus.Success) return hovered;
        if (!TryRefreshPointerTarget(element, out layout, out HtmlAutomationResult? failure)) return failure!;
        bool pointerDown = DispatchMouse(element, "pointerdown", true, true, layout, null);
        if (!TryRefreshPointerTarget(element, out layout, out failure)) return failure!;
        bool mouseDown = pointerDown && DispatchMouse(element, "mousedown", true, true, layout, null);
        if (!TryRefreshPointerTarget(element, out layout, out failure)) return failure!;
        if (mouseDown && RuntimeFocusController.CanFocus(element) && !_focus.Focus(element))
            return Failure(HtmlAutomationStatus.Rejected, "Page handlers redirected focus.", 1, Inspect(element));
        if (!TryRefreshPointerTarget(element, out layout, out failure)) return failure!;
        DispatchMouse(element, "pointerup", true, true, layout, null);
        if (!TryRefreshPointerTarget(element, out layout, out failure)) return failure!;
        if (pointerDown) DispatchMouse(element, "mouseup", true, true, layout, null);
        if (!TryRefreshPointerTarget(element, out layout, out failure)) return failure!;
        return Activate(element, focusTarget: false, layout);
    }

    private bool TryRefreshPointerTarget(IElement element, out RuntimeElementLayout layout, out HtmlAutomationResult? failure) {
        layout = _viewport.Measure(element, _viewport.CurrentCommandToken);
        if (!RuntimeFocusController.IsConnected(element) || RuntimeFocusController.Disabled(element)
            || !layout.IsVisible || !layout.IsInViewport || !layout.AcceptsPointerEvents || !layout.ReceivesPointerAtCenter) {
            if (ReferenceEquals(_hovered, element)) _hovered = null;
            failure = Failure(HtmlAutomationStatus.Rejected, "Pointer handlers changed the target's actionability.", 1, Inspect(element, layout));
            return false;
        }
        failure = null;
        return true;
    }

    private bool DispatchMouse(IElement element, string type, bool bubbles, bool cancelable,
        RuntimeElementLayout layout, IEventTarget? relatedTarget) {
        int x = (int)Math.Round((layout.Box?.X ?? 0D) + (layout.Box?.Width ?? 0D) / 2D);
        int y = (int)Math.Round((layout.Box?.Y ?? 0D) + (layout.Box?.Height ?? 0D) / 2D);
        var mouse = new MouseEvent(type, bubbles, cancelable, _document.DefaultView, 0,
            x, y, x, y, false, false, false, false, MouseButton.Primary, relatedTarget);
        element.Dispatch(mouse);
        return !mouse.IsDefaultPrevented;
    }

    private HtmlAutomationResult Press(IElement element, string requestedKey, HtmlKeyboardModifiers modifiers) {
        if (!TryNormalizeKey(requestedKey, out string key, out int keyCode))
            return Failure(HtmlAutomationStatus.InvalidValue, "Press supports one text element or Enter, Space, Tab, Escape, Backspace, Delete, Home, End and arrow keys.", 1, Inspect(element));
        if (!RuntimeFocusController.CanFocus(element))
            return Failure(HtmlAutomationStatus.Unsupported, "The target cannot receive keyboard focus.", 1, Inspect(element));
        bool wasFocused = ReferenceEquals(_focus.Focused, element);
        if (!_focus.Focus(element) || !ReferenceEquals(_focus.Focused, element))
            return Failure(HtmlAutomationStatus.Rejected, "Page handlers redirected focus.", 1, Inspect(element));
        if (!wasFocused && SupportsFill(element) && SelectionStart(element) == 0 && SelectionEnd(element) == 0)
            CollapseSelectionAtEnd(element);

        string modifierList = ModifierList(modifiers);
        var down = new KeyboardEvent("keydown", true, true, _document.DefaultView, keyCode, key,
            KeyboardLocation.Standard, modifierList, false);
        element.Dispatch(down);
        bool activateOnKeyUp = key == " " && SupportsSpaceActivation(element);
        HtmlAutomationResult result = down.IsDefaultPrevented || activateOnKeyUp ? Success(element) : ApplyKeyDefault(element, key, modifiers);
        IElement upTarget = _focus.Focused ?? element;
        var up = new KeyboardEvent("keyup", true, true, _document.DefaultView, keyCode, key,
            KeyboardLocation.Standard, modifierList, false);
        upTarget.Dispatch(up);
        if (activateOnKeyUp && !down.IsDefaultPrevented && !up.IsDefaultPrevented) {
            if (!ReferenceEquals(_focus.Focused, element) || !RuntimeFocusController.IsConnected(element)
                || RuntimeFocusController.Disabled(element) || RuntimeFocusController.HiddenByMarkup(element)
                || !SupportsSpaceActivation(element))
                return Failure(HtmlAutomationStatus.Rejected, "Keyboard handlers changed the activation target.", 1, Inspect(element));
            RuntimeElementLayout layout = _viewport.Measure(element, _viewport.CurrentCommandToken);
            if (!layout.IsVisible || !layout.IsInViewport)
                return Failure(HtmlAutomationStatus.Rejected, "Keyboard handlers changed the target's actionability.", 1, Inspect(element, layout));
            return Activate(element, focusTarget: false, modifiers: modifiers);
        }
        return result;
    }

    private HtmlAutomationResult ApplyKeyDefault(IElement element, string key, HtmlKeyboardModifiers modifiers) {
        if (SupportsFill(element)) {
            if (IsSelectAllShortcut(key, modifiers)) return SelectAllText(element);
            bool commandModifier = (modifiers & (HtmlKeyboardModifiers.Alt | HtmlKeyboardModifiers.Control | HtmlKeyboardModifiers.Meta)) != 0;
            if (!commandModifier && key == "Enter" && element is IHtmlInputElement && HtmlFormControlSemantics.ResolveFormOwner(element) is IHtmlFormElement form) {
                IHtmlElement? submitter = form.QuerySelectorAll("button,input").OfType<IHtmlElement>()
                    .FirstOrDefault(candidate => HtmlFormControlSemantics.IsSubmitter(candidate) && !RuntimeFocusController.Disabled(candidate));
                RuntimeFormActionOutcome outcome = SubmitForm(form, submitter, dispatchSubmitEvent: true, validate: true);
                return outcome.Status == HtmlAutomationStatus.Success ? Success(element)
                    : Failure(outcome.Status, outcome.Message!, 1, Inspect(element));
            }
            if (!commandModifier && key is "Home" or "End" or "ArrowLeft" or "ArrowRight")
                return MoveTextCaret(element, key, modifiers.HasFlag(HtmlKeyboardModifiers.Shift));
            if (!commandModifier && key is "Backspace" or "Delete") return EditSelection(element, key, string.Empty);
            if (!commandModifier && key == "Enter" && element is IHtmlTextAreaElement) return EditSelection(element, key, "\n");
            if (!commandModifier && IsTextKey(key)) return EditSelection(element, key, key);
        }
        if (key == "Tab" && (modifiers & (HtmlKeyboardModifiers.Alt | HtmlKeyboardModifiers.Control | HtmlKeyboardModifiers.Meta)) == 0)
            return MoveFocus(element, modifiers.HasFlag(HtmlKeyboardModifiers.Shift));
        bool enter = key == "Enter" && (element is IHtmlButtonElement or IHtmlAnchorElement
            || element is IHtmlInputElement input && input.Type is "button" or "submit" or "reset" or "image");
        bool space = key == " " && (element is IHtmlButtonElement
            || element is IHtmlInputElement choice && choice.Type is "button" or "submit" or "reset" or "checkbox" or "radio");
        return enter || space ? Activate(element, focusTarget: false, modifiers: modifiers) : Success(element);
    }

    private HtmlAutomationResult SetSelection(IElement element, int start, int end) {
        if (!SupportsFill(element))
            return Failure(HtmlAutomationStatus.Unsupported, "SetSelection supports editable text inputs and textareas.", 1, Inspect(element));
        if (!ReadyToEdit(element))
            return Failure(HtmlAutomationStatus.NotReady, "The text control is not editable.", 1, Inspect(element));
        if (!_focus.Focus(element) || !ReferenceEquals(_focus.Focused, element) || !SupportsFill(element) || !ReadyToEdit(element))
            return Failure(HtmlAutomationStatus.Rejected, "Focus handlers changed the selection target.", 1, Inspect(element));
        SetTextSelection(element, start, end, "forward");
        return Success(element);
    }

    private HtmlAutomationResult SelectAllText(IElement element) {
        string current = RuntimeFocusController.Value(element) ?? string.Empty;
        SetTextSelection(element, 0, current.Length, "forward");
        return Success(element);
    }

    private HtmlAutomationResult EditSelection(IElement element, string key, string inserted) {
        string current = RuntimeFocusController.Value(element) ?? string.Empty;
        TextSelection selection = GetTextSelection(element, current.Length);
        int replaceStart = selection.Start;
        int replaceEnd = selection.End;
        if (replaceStart == replaceEnd && key == "Backspace" && replaceStart > 0) {
            replaceStart = PreviousTextElementBoundary(current, replaceStart);
            replaceEnd = NextTextElementBoundary(current, replaceStart);
        }
        if (replaceStart == replaceEnd && key == "Delete" && replaceEnd < current.Length) {
            replaceStart = TextElementStart(current, replaceStart);
            replaceEnd = NextTextElementBoundary(current, replaceStart);
        }
        string next = current[..replaceStart] + inserted + current[replaceEnd..];
        if (next == current) return Success(element);
        var before = new InputEvent("beforeinput", true, true, inserted);
        element.Dispatch(before);
        if (before.IsDefaultPrevented) return Success(element);
        TextSelection afterHandler = GetTextSelection(element, current.Length);
        if (!ReferenceEquals(_focus.Focused, element) || !ReadyToEdit(element)
            || RuntimeFocusController.Value(element) != current || afterHandler.Start != selection.Start || afterHandler.End != selection.End)
            return Failure(HtmlAutomationStatus.Rejected, "Beforeinput handlers changed the editing target.", 1, Inspect(element));
        _focus.ChangedByUser(element);
        if (element is IHtmlInputElement input) input.Value = next;
        else ((IHtmlTextAreaElement)element).Value = next;
        int caret = replaceStart + inserted.Length;
        SetTextSelection(element, caret, caret, "none");
        element.Dispatch(new InputEvent("input", true, false, inserted));
        return Success(element);
    }

    private HtmlAutomationResult MoveTextCaret(IElement element, string key, bool extend) {
        string value = RuntimeFocusController.Value(element) ?? string.Empty;
        TextSelection selection = GetTextSelection(element, value.Length);
        if (!extend) {
            int collapsed = key switch {
                "Home" => 0,
                "End" => value.Length,
                "ArrowLeft" => selection.Start != selection.End ? selection.Start : PreviousTextElementBoundary(value, selection.Start),
                _ => selection.Start != selection.End ? selection.End : NextTextElementBoundary(value, selection.End)
            };
            SetTextSelection(element, collapsed, collapsed, "none");
            return Success(element);
        }
        int anchor = selection.Direction == "backward" ? selection.End : selection.Start;
        int focus = selection.Direction == "backward" ? selection.Start : selection.End;
        int moved = key switch {
            "Home" => 0,
            "End" => value.Length,
            "ArrowLeft" => PreviousTextElementBoundary(value, focus),
            _ => NextTextElementBoundary(value, focus)
        };
        SetTextSelection(element, Math.Min(anchor, moved), Math.Max(anchor, moved), moved < anchor ? "backward" : "forward");
        return Success(element);
    }

    private HtmlAutomationResult MoveFocus(IElement element, bool reverse) {
        IElement[] candidates = _document.QuerySelectorAll("*")
            .Select((candidate, index) => new { Element = candidate, Index = index, TabIndex = TabIndex(candidate) })
            .Where(candidate => candidate.TabIndex >= 0 && RuntimeFocusController.CanFocus(candidate.Element))
            .OrderBy(candidate => candidate.TabIndex == 0 ? int.MaxValue : candidate.TabIndex)
            .ThenBy(candidate => candidate.Index)
            .Select(candidate => candidate.Element)
            .ToArray();
        int current = Array.IndexOf(candidates, element);
        int nextIndex = current < 0 ? reverse ? candidates.Length - 1 : 0
            : (current + (reverse ? -1 : 1) + candidates.Length) % candidates.Length;
        IElement? next = candidates.Length == 0 ? null : candidates[nextIndex];
        return next != null && _focus.Focus(next) ? Success(element)
            : Failure(HtmlAutomationStatus.Rejected, "Tab focus traversal was redirected.", 1, Inspect(element));
    }

    private static int? SelectionStart(IElement element) => element switch {
        IHtmlInputElement input when SupportsFill(input) => input.SelectionStart,
        IHtmlTextAreaElement area => area.SelectionStart,
        _ => null
    };

    private static int? SelectionEnd(IElement element) => element switch {
        IHtmlInputElement input when SupportsFill(input) => input.SelectionEnd,
        IHtmlTextAreaElement area => area.SelectionEnd,
        _ => null
    };

    private static TextSelection GetTextSelection(IElement element, int length) => element switch {
        IHtmlInputElement input => new(Math.Clamp(input.SelectionStart, 0, length), Math.Clamp(input.SelectionEnd, 0, length), input.SelectionDirection ?? "none"),
        IHtmlTextAreaElement area => new(Math.Clamp(area.SelectionStart, 0, length), Math.Clamp(area.SelectionEnd, 0, length), area.SelectionDirection ?? "none"),
        _ => new(0, 0, "none")
    };

    private static void SetTextSelection(IElement element, int start, int end, string direction) {
        if (element is IHtmlInputElement input) input.Select(start, end, direction);
        else if (element is IHtmlTextAreaElement area) area.Select(start, end, direction);
    }

    private static void CollapseSelectionAtEnd(IElement element) {
        int end = (RuntimeFocusController.Value(element) ?? string.Empty).Length;
        SetTextSelection(element, end, end, "none");
    }

    private static int PreviousTextElementBoundary(string value, int offset) {
        int previous = 0;
        foreach (int boundary in StringInfo.ParseCombiningCharacters(value)) {
            if (boundary >= offset) break;
            previous = boundary;
        }
        return previous;
    }

    private static int NextTextElementBoundary(string value, int offset) {
        foreach (int boundary in StringInfo.ParseCombiningCharacters(value))
            if (boundary > offset) return boundary;
        return value.Length;
    }

    private static int TextElementStart(string value, int offset) {
        int start = 0;
        foreach (int boundary in StringInfo.ParseCombiningCharacters(value)) {
            if (boundary > offset) break;
            start = boundary;
        }
        return start;
    }

    private static bool IsSelectAllShortcut(string key, HtmlKeyboardModifiers modifiers) =>
        key.Equals("a", StringComparison.OrdinalIgnoreCase)
        && (modifiers.HasFlag(HtmlKeyboardModifiers.Control) ^ modifiers.HasFlag(HtmlKeyboardModifiers.Meta))
        && !modifiers.HasFlag(HtmlKeyboardModifiers.Alt);

    private static string ModifierList(HtmlKeyboardModifiers modifiers) => string.Join(' ', new[] {
        modifiers.HasFlag(HtmlKeyboardModifiers.Alt) ? "Alt" : null,
        modifiers.HasFlag(HtmlKeyboardModifiers.Control) ? "Control" : null,
        modifiers.HasFlag(HtmlKeyboardModifiers.Meta) ? "Meta" : null,
        modifiers.HasFlag(HtmlKeyboardModifiers.Shift) ? "Shift" : null
    }.Where(value => value != null));

    private static int TabIndex(IElement element) {
        string? value = element.GetAttribute("tabindex");
        if (value != null && int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)) return parsed;
        return element is IHtmlTextAreaElement or IHtmlSelectElement or IHtmlButtonElement
            || element is IHtmlInputElement input && input.Type != "hidden"
            || element is IHtmlAnchorElement && element.HasAttribute("href") ? 0 : -1;
    }

    private static bool TryNormalizeKey(string requested, out string key, out int keyCode) {
        key = requested == "Space" ? " " : requested;
        keyCode = key switch {
            "Backspace" => 8, "Tab" => 9, "Enter" => 13, "Escape" => 27, " " => 32,
            "End" => 35, "Home" => 36, "ArrowLeft" => 37, "ArrowUp" => 38,
            "ArrowRight" => 39, "ArrowDown" => 40, "Delete" => 46, _ => 0
        };
        if (keyCode != 0) return true;
        var enumerator = StringInfo.GetTextElementEnumerator(key);
        if (!enumerator.MoveNext() || enumerator.GetTextElement() != key || enumerator.MoveNext()) return false;
        keyCode = char.ConvertToUtf32(key, 0);
        return !char.IsControl(key, 0);
    }

    private static bool IsTextKey(string key) => !string.IsNullOrEmpty(key)
        && key is not ("Enter" or "Tab" or "Escape" or "Backspace" or "Delete" or "Home" or "End"
            or "ArrowLeft" or "ArrowUp" or "ArrowRight" or "ArrowDown");

    private static bool SupportsSpaceActivation(IElement element) => element is IHtmlButtonElement
        || element is IHtmlInputElement input && input.Type is "button" or "submit" or "reset" or "checkbox" or "radio";

    private readonly record struct TextSelection(int Start, int End, string Direction);
}
