namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
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
