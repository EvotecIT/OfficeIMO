namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    private readonly Dictionary<IElement, bool> _specializedIntrinsicWidthContent = new Dictionary<IElement, bool>();

    private void CaptureIntrinsicWidths(
        IElement element, HtmlComputedStyle computed, HtmlRenderBoxStyle style, HtmlRenderBoxStyle? parent,
        bool pseudoElement, double reference, double fontSize) {
        HtmlRenderIntrinsicWidth? width = ReadIntrinsicWidth(computed.GetValue("width"), reference, fontSize);
        HtmlRenderIntrinsicWidth? minimum = ReadIntrinsicWidth(computed.GetValue("min-width"), reference, fontSize);
        HtmlRenderIntrinsicWidth? maximum = ReadIntrinsicWidth(computed.GetValue("max-width"), reference, fontSize);
        if (!width.HasValue && !minimum.HasValue && !maximum.HasValue) return;
        if (pseudoElement || style.Display is not ("block" or "inline-block") || style.Position != "static"
            || style.FloatSide != "none" || style.WritingMode != "horizontal-tb"
            || style.ColumnCount.HasValue || style.ColumnWidth.HasValue || style.ContainerType != "normal"
            || parent?.WritingMode is "vertical-rl" or "vertical-lr") return;
        if (element.LocalName.ToLowerInvariant() is "html" or "body" or "table" or "tr" or "td" or "th"
            or "img" or "svg" or "iframe" or "input" or "textarea" or "select" or "button"
            or "progress" or "meter" or "math" or "ruby" or "hr") return;

        // Use the original formatting context, before allocated flex/grid items
        // and atomic inlines are blockified by their specialized owners.
        string parentDisplay = parent?.Display ?? string.Empty;
        IElement? ancestor = element.ParentElement;
        while (parentDisplay == "contents" && ancestor?.ParentElement != null) {
            ancestor = ancestor.ParentElement;
            parentDisplay = _computedStyles.Elements.TryGetValue(ancestor, out HtmlComputedStyle? ancestorStyle)
                ? ResolveDisplay(ancestor, ancestorStyle.GetValue("display")) : HtmlElementDisplay.GetDefaultValue(ancestor);
        }
        if (parentDisplay is "flex" or "inline-flex" or "grid" or "inline-grid") return;
        if (HasSpecializedIntrinsicWidthContent(element)) return;
        style.IntrinsicWidth = width;
        style.IntrinsicMinWidth = minimum;
        style.IntrinsicMaxWidth = maximum;
    }

    private HtmlRenderIntrinsicWidth? ReadIntrinsicWidth(string value, double reference, double fontSize) {
        string normalized = OfficeIMO.Html.Css.HtmlCssTokenizer.StripComments(value).Trim().ToLowerInvariant();
        if (normalized == "min-content") return new(HtmlRenderIntrinsicWidthKind.MinContent);
        if (normalized == "max-content") return new(HtmlRenderIntrinsicWidthKind.MaxContent);
        if (normalized == "fit-content") return new(HtmlRenderIntrinsicWidthKind.FitContent);
        if (!normalized.StartsWith("fit-content(", StringComparison.Ordinal) || !normalized.EndsWith(")", StringComparison.Ordinal)) return null;
        string argument = normalized.Substring(12, normalized.Length - 13);
        // Percentage functions require separate cyclic-reference qualification.
        if (argument.IndexOf('%') >= 0) return null;
        double? limit = ReadLength(argument, null, reference, fontSize);
        return limit.HasValue ? new(HtmlRenderIntrinsicWidthKind.FitContent, limit.Value) : null;
    }

    /// <summary>Removes cyclic references only for an ordinary descendant's intrinsic contribution.</summary>
    internal HtmlRenderBoxStyle ResolveIntrinsicMeasurementStyle(IElement element, HtmlRenderBoxStyle style) {
        _computedStyles.Elements.TryGetValue(element, out HtmlComputedStyle? computed);
        HtmlComputedStyle? physical = computed == null ? null : PhysicalizeLogicalProperties(computed, style.WritingMode, style.Direction);
        string leftPadding = physical?.GetValue("padding-left") ?? string.Empty;
        string rightPadding = physical?.GetValue("padding-right") ?? string.Empty;
        string leftMargin = physical?.GetValue("margin-left") ?? string.Empty;
        string rightMargin = physical?.GetValue("margin-right") ?? string.Empty;
        bool Cyclic(string value) => value.IndexOf('%') >= 0;
        if (!style.ExplicitWidthUsesPercentage && !style.MaxWidthUsesPercentage && !style.MinWidthWithIndefiniteReference.HasValue
            && !Cyclic(leftPadding) && !Cyclic(rightPadding) && !Cyclic(leftMargin) && !Cyclic(rightMargin)) return style;
        HtmlRenderBoxStyle resolved = style.Clone();
        // Preferred and maximum percentages use their initial values; minimum
        // percentages and box edges retain their absolute part against zero.
        if (style.ExplicitWidthUsesPercentage) {
            resolved.ExplicitWidth = null;
            resolved.ExplicitWidthUsesPercentage = false;
        }
        if (style.MaxWidthUsesPercentage) {
            resolved.MaxWidth = null;
            resolved.MaxWidthUsesPercentage = false;
        }
        if (style.MinWidthWithIndefiniteReference.HasValue) resolved.MinWidth = style.MinWidthWithIndefiniteReference;
        if (Cyclic(leftPadding)) resolved.PaddingLeft = Math.Max(0D, ReadLength(leftPadding, null, 0D, style.Font.Size) ?? 0D);
        if (Cyclic(rightPadding)) resolved.PaddingRight = Math.Max(0D, ReadLength(rightPadding, null, 0D, style.Font.Size) ?? 0D);
        if (Cyclic(leftMargin)) ApplyMarginLength(leftMargin, 0D, style.Font.Size, ref resolved.MarginLeft);
        if (Cyclic(rightMargin)) ApplyMarginLength(rightMargin, 0D, style.Font.Size, ref resolved.MarginRight);
        return resolved;
    }

    private bool HasSpecializedIntrinsicWidthContent(IElement element) {
        if (_specializedIntrinsicWidthContent.TryGetValue(element, out bool specialized)) return specialized;
        var pending = new Stack<IElement>(element.Children);
        while (pending.Count > 0) {
            _cancellationToken.ThrowIfCancellationRequested();
            IElement child = pending.Pop();
            _computedStyles.Elements.TryGetValue(child, out HtmlComputedStyle? computed);
            string display = ResolveDisplay(child, computed?.GetValue("display") ?? string.Empty);
            string position = computed?.GetValue("position") ?? string.Empty;
            if (display == "none" || position is "absolute" or "fixed") continue;
            string writingMode = computed?.GetValue("writing-mode") ?? string.Empty;
            string floating = computed?.GetValue("float") ?? string.Empty;
            string columns = computed?.GetValue("column-count") ?? string.Empty;
            string columnWidth = computed?.GetValue("column-width") ?? string.Empty;
            string container = computed?.GetValue("container-type") ?? string.Empty;
            if (child.LocalName is "table" or "math" or "ruby" or "img" or "svg" or "iframe"
                or "input" or "textarea" or "select" or "button" or "progress" or "meter"
                || display is "table" or "inline-table" or "flex" or "inline-flex" or "grid" or "inline-grid"
                || display.StartsWith("table-", StringComparison.Ordinal)
                || writingMode is "vertical-rl" or "vertical-lr"
                || floating.Length > 0 && floating != "none"
                || columns.Length > 0 && columns != "auto"
                || columnWidth.Length > 0 && columnWidth != "auto"
                || container.Length > 0 && container != "normal"
                || display == "inline" && HasUnqualifiedInlineIntrinsicEdges(computed)) {
                _specializedIntrinsicWidthContent[element] = true;
                return true;
            }
            foreach (IElement descendant in child.Children) pending.Push(descendant);
        }
        _specializedIntrinsicWidthContent[element] = false;
        return false;
    }

    private bool HasUnqualifiedInlineIntrinsicEdges(HtmlComputedStyle? computed) {
        if (computed == null) return false;
        HtmlComputedStyle physical = PhysicalizeLogicalProperties(computed, "horizontal-tb",
            computed.GetValue("direction") == "rtl" ? "rtl" : "ltr");
        // Inline decoration paint expands around text, but line layout does not
        // reserve these edges yet. Measuring them alone can cause false wrapping.
        foreach (string property in new[] { "margin-left", "margin-right", "padding-left", "padding-right" }) {
            string value = physical.GetValue(property).Trim();
            if (value.IndexOf('%') >= 0 && value != "0%") return true;
        }
        var edges = new HtmlRenderBoxStyle();
        ApplyBoxValues(physical, 1D, _rootFontSize, edges);
        return edges.HorizontalInsets > 0D || edges.MarginLeft != 0D || edges.MarginRight != 0D;
    }
}
