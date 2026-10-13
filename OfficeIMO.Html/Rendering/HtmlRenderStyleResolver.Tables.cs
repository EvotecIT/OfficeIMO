namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    private string ResolveTableInternalDimension(string value, HtmlRenderBoxStyle style) {
        if (style.Display is not ("table-column" or "table-column-group" or "table-row" or "table-row-group"
            or "table-header-group" or "table-footer-group" or "table-cell")) return value;
        return HtmlRenderCssValues.IsTableInternalAutoMath(value, style.Font.Size, _rootFontSize,
            _viewportWidth, _viewportHeight, style.ContainerUnitWidth, style.ContainerUnitHeight, style.CharacterAdvance,
            (style.WritingMode is "vertical-rl" or "vertical-lr") && style.TextOrientation == "upright") ? "auto" : value;
    }

    private static void CaptureTablePercentageHeight(string cssHeight, string? attributeHeight, HtmlRenderBoxStyle style) {
        if (style.Display is not ("table-row" or "table-cell" or "table-row-group" or "table-header-group" or "table-footer-group")) return;
        string height = cssHeight.Length > 0 ? cssHeight : NormalizeHtmlDimensionAttribute(attributeHeight);
        if (height.IndexOf('%') >= 0) style.TablePercentageHeight = height;
    }

    /// <summary>Uses the shared computed-length resolver with the table owner's final percentage basis.</summary>
    internal double? ResolveTablePercentageHeight(HtmlRenderBoxStyle style, double reference, string? effectiveValue = null) {
        if (!HtmlRenderCssValues.TryLength(effectiveValue ?? style.TablePercentageHeight, reference, style.Font.Size, _rootFontSize,
                _viewportWidth, _viewportHeight, style.ContainerUnitWidth ?? double.NaN, style.ContainerUnitHeight ?? double.NaN,
                out double height, out bool calculated, style.CharacterAdvance,
                (style.WritingMode is "vertical-rl" or "vertical-lr") && style.TextOrientation == "upright")) return null;
        return height >= 0D ? height : calculated ? 0D : null;
    }

    private void ApplyTableDefaults(IElement element, HtmlComputedStyle computed, HtmlRenderBoxStyle? parent, HtmlRenderBoxStyle style) {
        string tag = element.LocalName.ToLowerInvariant();
        if (tag == "table" && !HasAuthoredValue(computed, "box-sizing")) style.BorderBox = true;
        string alignment = computed.GetValue("vertical-align").Trim().ToLowerInvariant();
        style.TableVerticalAlignment = computed.IsInheritedValue("vertical-align")
            ? parent?.TableVerticalAlignment ?? "baseline"
            : alignment.Length == 0 || computed.IsResetValue("vertical-align") ? "baseline" : alignment;
        if (_options.UserAgentStyles == HtmlRenderUserAgentStyleMode.Browser
            && !HasAuthoredValue(computed, "vertical-align")) {
            // HTML row groups default to middle; rows and cells inherit the
            // group or row value through their user-agent declarations.
            if (tag == "thead" || tag == "tbody" || tag == "tfoot") style.TableVerticalAlignment = "middle";
            else if (tag == "tr" || IsTableCellElement(element)) style.TableVerticalAlignment = parent?.TableVerticalAlignment ?? "middle";
        }
        if (!IsTableCellElement(element)) return;

        // Apply defaults before the authored box values. Zero, initial, inherit,
        // and logical declarations must not be mistaken for a missing value.
        if (!HasAuthoredValue(computed, "padding-top")) style.PaddingTop = 2D;
        if (!HasAuthoredValue(computed, "padding-right")) style.PaddingRight = 2D;
        if (!HasAuthoredValue(computed, "padding-bottom")) style.PaddingBottom = 2D;
        if (!HasAuthoredValue(computed, "padding-left")) style.PaddingLeft = 2D;
    }

    private static bool IsTableCellElement(IElement element) =>
        element.LocalName.Equals("td", StringComparison.OrdinalIgnoreCase)
        || element.LocalName.Equals("th", StringComparison.OrdinalIgnoreCase);
}
