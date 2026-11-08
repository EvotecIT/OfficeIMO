using AngleSharp.Dom;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;

namespace OfficeIMO.Word.Html;

internal partial class HtmlToWordConverter {
    /// <summary>Replays effective values through the importer's existing paint and border mappings.</summary>
    private static void LowerMappedLayerValues(IElement element, Dictionary<string, string> properties) {
        if (properties.ContainsKey("background-color")) properties.Remove("background");
        if (properties.TryGetValue("background-image", out string? image) && image.Equals("none", StringComparison.OrdinalIgnoreCase)) {
            foreach (var property in properties.Where(pair => pair.Key.StartsWith("background-", StringComparison.OrdinalIgnoreCase)
                         && !pair.Key.Equals("background-color", StringComparison.OrdinalIgnoreCase)).ToArray()) {
                string initial = HtmlComputedStyleEngine.BackgroundInitialValue(property.Key);
                if (!string.IsNullOrEmpty(initial) && property.Value.Equals(initial, StringComparison.OrdinalIgnoreCase)) properties.Remove(property.Key);
            }
        }

        if (element.LocalName is not ("table" or "tr" or "td" or "th")) return;
        foreach (string side in new[] { "top", "right", "bottom", "left" }) {
            string prefix = "border-" + side;
            if (!properties.TryGetValue(prefix + "-style", out string? style)) continue;
            // A nonpainting side has no effective width or colour to approximate.
            string border;
            if (style.Equals("none", StringComparison.OrdinalIgnoreCase)) {
                border = "none";
            } else {
                if (!properties.TryGetValue(prefix + "-width", out string? width) ||
                    !properties.TryGetValue(prefix + "-color", out string? color)) continue;
                border = width + " " + style + " " + color;
            }
            if (!TryParseBorder(border, out _, out _, out _)) continue;
            properties[prefix] = border;
            properties.Remove(prefix + "-width");
            properties.Remove(prefix + "-style");
            properties.Remove(prefix + "-color");
        }
    }

    /// <summary>Retains source rhythm for bare text whose native paragraph is created by the importer.</summary>
    private void ApplyLayeredParagraphDefaults(WordParagraph paragraph, IElement? source) {
        if (!_useCanonicalLayerCascade) return;
        for (IElement? ancestor = source; ancestor != null; ancestor = ancestor.ParentElement) {
            if (ancestor.LocalName == "body" || IsBlockSpacingElement(ancestor.LocalName) || IsTableCellElement(ancestor.LocalName)) {
                ApplyParagraphLineHeight(paragraph, ParseElementBoxStyles(ancestor));
                return;
            }
        }
    }

    private static void ApplyParagraphLineHeight(WordParagraph paragraph, CssStyleMapper.CssProperties parsed) {
        if (!parsed.LineHeight.HasValue) return;
        // CSS parsing is native-shaped; the Word owner interprets its units.
        paragraph.LineSpacing = WordLineSpacing.Read(new SpacingBetweenLines {
            Line = parsed.LineHeight.Value.ToString(System.Globalization.CultureInfo.InvariantCulture),
            LineRule = parsed.LineHeightRule
        });
    }
}
