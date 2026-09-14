using AngleSharp.Dom;

namespace OfficeIMO.Html.Runtime.Worker;

internal sealed partial class RuntimeAutomation {
    internal HtmlPageObservation Observe(HtmlPageObservationRequest request, string contextId, string pageId, long revision, CancellationToken token) {
        IReadOnlyList<RuntimeDocumentElement> all = _locators.Enumerate(token);
        var elements = new List<HtmlObservedElement>(Math.Min(request.MaxElements, all.Count));
        var includedIndices = new Dictionary<int, int>();
        int textCharacters = 0;
        bool truncated = false;
        RuntimeElementLayout documentLayout = _document.DocumentElement == null
            ? RuntimeElementLayout.Unqualified
            : _viewport.Measure(_document.DocumentElement, token);

        for (int index = 0; index < all.Count; index++) {
            token.ThrowIfCancellationRequested();
            RuntimeDocumentElement item = all[index];
            HtmlRuntimeElementState state = Inspect(item.Element, token);
            bool actionable = IsActionable(item.Element, state);
            if (!request.IncludeHidden && (state.IsHiddenByMarkup || state.IsVisible == false)) continue;
            if (request.ActionableOnly && !actionable) continue;
            if (elements.Count >= request.MaxElements) { truncated = true; break; }

            string text = request.Mode == HtmlPageObservationMode.Visual ? string.Empty : state.Text;
            string name = request.Mode == HtmlPageObservationMode.Visual ? string.Empty : state.AccessibleName;
            int remaining = request.MaxTextCharacters - textCharacters;
            if (remaining <= 0 && (text.Length > 0 || name.Length > 0)) { text = string.Empty; name = string.Empty; truncated = true; }
            else {
                name = Limit(name, ref remaining, ref truncated);
                text = Limit(text, ref remaining, ref truncated);
                textCharacters = request.MaxTextCharacters - remaining;
            }
            bool includeVisual = request.Mode != HtmlPageObservationMode.Semantic;
            int? parentElementIndex = item.ParentElementIndex;
            while (parentElementIndex != null && !includedIndices.TryGetValue(parentElementIndex.Value, out _))
                parentElementIndex = all[parentElementIndex.Value].ParentElementIndex;
            int? observedParentIndex = parentElementIndex == null ? null : includedIndices[parentElementIndex.Value];
            includedIndices[index] = elements.Count;
            elements.Add(new HtmlObservedElement {
                Reference = new HtmlObservedElementReference { PageId = pageId, Revision = revision, ElementIndex = index,
                    ElementName = item.Element.LocalName, ElementId = item.Element.Id ?? string.Empty },
                Depth = item.Depth,
                ParentElementIndex = observedParentIndex,
                ElementName = item.Element.LocalName,
                Role = request.Mode == HtmlPageObservationMode.Visual ? string.Empty : Role(item.Element),
                AccessibleName = name,
                Text = text,
                Value = request.Mode == HtmlPageObservationMode.Visual ? null : state.Value,
                SelectedValues = request.Mode == HtmlPageObservationMode.Visual ? Array.Empty<string>() : state.SelectedValues,
                SelectionStart = request.Mode == HtmlPageObservationMode.Visual ? null : state.SelectionStart,
                SelectionEnd = request.Mode == HtmlPageObservationMode.Visual ? null : state.SelectionEnd,
                IsChecked = request.Mode == HtmlPageObservationMode.Visual ? null : state.IsChecked,
                IsDisabled = state.IsDisabled,
                IsEditable = state.IsEditable,
                IsFocused = state.IsFocused,
                IsVisible = includeVisual ? state.IsVisible : null,
                IsInViewport = includeVisual ? state.IsInViewport : null,
                IsActionable = actionable,
                BoundingBox = includeVisual ? state.BoundingBox : null
            });
        }

        return new HtmlPageObservation {
            ProviderId = HtmlRuntimeProviderDescriptor.ProcessWorker.Id,
            ContextId = contextId,
            PageId = pageId,
            Revision = revision,
            Url = new Uri(_document.Url),
            Title = _document.Title ?? string.Empty,
            Mode = request.Mode,
            ViewportWidth = _viewport.Width,
            ViewportHeight = _viewport.Height,
            ScrollX = _viewport.ScrollX,
            ScrollY = _viewport.ScrollY,
            DocumentWidth = documentLayout.DocumentWidth,
            DocumentHeight = documentLayout.DocumentHeight,
            IsTruncated = truncated,
            Elements = Array.AsReadOnly(elements.ToArray()),
            Diagnostics = _viewport.Enabled ? Array.Empty<string>() : new[] { "Visual geometry requires WebApplicationV1." }
        };
    }

    private static string Limit(string value, ref int remaining, ref bool truncated) {
        if (value.Length <= remaining) { remaining -= value.Length; return value; }
        string result = remaining == 0 ? string.Empty : value[..remaining];
        remaining = 0;
        truncated = true;
        return result;
    }

    private static bool IsActionable(IElement element, HtmlRuntimeElementState state) {
        bool semantic = element.LocalName is "a" or "button" or "input" or "select" or "textarea" or "summary"
            || element.HasAttribute("tabindex") || element.GetAttribute("role") is "button" or "link" or "checkbox" or "radio" or "textbox" or "combobox";
        return semantic && !state.IsDisabled && !state.IsHiddenByMarkup && state.IsConnected && state.IsVisible != false;
    }

    private static string Role(IElement element) {
        string? explicitRole = element.GetAttribute("role")?.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
        if (!string.IsNullOrEmpty(explicitRole)) return explicitRole;
        return element.LocalName switch {
            "a" when element.HasAttribute("href") => "link",
            "button" => "button",
            "textarea" => "textbox",
            "select" => "combobox",
            "option" => "option",
            "img" => "img",
            "table" => "table",
            "tr" => "row",
            "td" => "cell",
            "th" => "columnheader",
            "h1" or "h2" or "h3" or "h4" or "h5" or "h6" => "heading",
            "input" => element.GetAttribute("type")?.ToLowerInvariant() switch {
                "checkbox" => "checkbox", "radio" => "radio", "button" or "submit" or "reset" => "button", _ => "textbox"
            },
            _ => string.Empty
        };
    }
}
