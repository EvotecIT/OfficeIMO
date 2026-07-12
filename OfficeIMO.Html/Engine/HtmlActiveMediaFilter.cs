using AngleSharp.Css.Dom;
using AngleSharp.Css.Parser;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using System.Text;

namespace OfficeIMO.Html;

/// <summary>
/// Filters shared HTML adapter input to the active CSS media context.
/// </summary>
public static class HtmlActiveMediaFilter {
    private const string FragmentRootElementName = "officeimo-fragment-root";

    /// <summary>
    /// Removes inactive stylesheet links, picture sources, and media-gated style rules for the requested media context.
    /// </summary>
    /// <param name="html">HTML document or fragment to filter.</param>
    /// <param name="mediaContext">CSS media context used by the target conversion profile.</param>
    /// <returns>Filtered HTML when changes were required; otherwise the original HTML.</returns>
    public static string Filter(string html, HtmlCssMediaContext mediaContext) {
        if (string.IsNullOrWhiteSpace(html)) {
            return html;
        }

        try {
            bool isFragment = !ContainsHtmlDocumentElement(html);
            IHtmlDocument parsed = HtmlDocumentParser.ParseDocument(isFragment
                ? "<html><body><" + FragmentRootElementName + ">" + html + "</" + FragmentRootElementName + "></body></html>"
                : html);
            bool changed = FilterDocument(parsed, mediaContext);

            if (!changed) {
                return html;
            }

            if (isFragment) {
                return parsed.QuerySelector(FragmentRootElementName)?.InnerHtml ?? string.Empty;
            }

            return parsed.DocumentElement?.OuterHtml ?? html;
        } catch {
            return html;
        }
    }

    /// <summary>
    /// Removes inactive media content from a prepared DOM in place without serializing and reparsing it.
    /// </summary>
    /// <returns>Whether the document was changed.</returns>
    public static bool Filter(IHtmlDocument document, HtmlCssMediaContext mediaContext) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return FilterDocument(document, mediaContext);
    }

    private static bool FilterDocument(IHtmlDocument parsed, HtmlCssMediaContext mediaContext) {
        bool changed = false;
        foreach (IHtmlLinkElement linkElement in parsed.QuerySelectorAll("link").OfType<IHtmlLinkElement>()) {
            if (string.Equals(linkElement.Relation, "stylesheet", StringComparison.OrdinalIgnoreCase)
                && !HtmlComputedStyleEngine.IsApplicableMedia(linkElement.GetAttribute("media") ?? string.Empty, mediaContext)) {
                linkElement.Remove();
                changed = true;
            }
        }

        foreach (IElement sourceElement in parsed.QuerySelectorAll("picture > source")) {
            if (!HtmlComputedStyleEngine.IsApplicableMedia(sourceElement.GetAttribute("media") ?? string.Empty, mediaContext)) {
                sourceElement.Remove();
                changed = true;
                continue;
            }

            if (!HtmlPictureSourceSupport.IsSupportedConversionContentType(sourceElement.GetAttribute("type"))) {
                sourceElement.Remove();
                changed = true;
            }
        }

        foreach (IHtmlStyleElement styleElement in parsed.QuerySelectorAll("style").OfType<IHtmlStyleElement>()) {
            if (!IsCssStyleElement(styleElement)
                || !HtmlComputedStyleEngine.IsApplicableMedia(styleElement.GetAttribute("media") ?? string.Empty, mediaContext)) {
                styleElement.Remove();
                changed = true;
                continue;
            }

            string expanded = ExpandActiveMediaStyleRules(styleElement.TextContent, mediaContext, out bool stylesheetChanged);
            if (stylesheetChanged) {
                styleElement.TextContent = expanded;
                changed = true;
            }
        }

        return changed;
    }

    private static bool ContainsHtmlDocumentElement(string html) {
        for (int i = 0; i < html.Length; i++) {
            if (html[i] != '<') {
                continue;
            }

            int cursor = i + 1;
            while (cursor < html.Length && char.IsWhiteSpace(html[cursor])) {
                cursor++;
            }

            if (cursor < html.Length && html[cursor] == '/') {
                cursor++;
                while (cursor < html.Length && char.IsWhiteSpace(html[cursor])) {
                    cursor++;
                }
            }

            if (cursor + 4 > html.Length || !html.Substring(cursor).StartsWith("html", StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            int end = cursor + 4;
            if (end == html.Length || char.IsWhiteSpace(html[end]) || html[end] == '>' || html[end] == '/') {
                return true;
            }
        }

        return false;
    }

    private static bool IsCssStyleElement(IHtmlStyleElement styleElement) {
        string type = styleElement.GetAttribute("type") ?? string.Empty;
        if (string.IsNullOrWhiteSpace(type)) {
            return true;
        }

        string mediaType = type.Split(';')[0].Trim();
        return string.Equals(mediaType, "text/css", StringComparison.OrdinalIgnoreCase);
    }

    private static string ExpandActiveMediaStyleRules(string css, HtmlCssMediaContext mediaContext, out bool changed) {
        changed = false;
        if (string.IsNullOrWhiteSpace(css)) {
            return css;
        }

        try {
            var parser = new CssParser();
            var stylesheet = parser.ParseStyleSheet(css);
            var builder = new StringBuilder();
            foreach (ICssRule rule in stylesheet.Rules) {
                AppendActiveMediaStyleRule(builder, rule, mediaContext, ref changed);
            }

            return changed ? builder.ToString() : css;
        } catch {
            return css;
        }
    }

    private static void AppendActiveMediaStyleRule(StringBuilder builder, ICssRule rule, HtmlCssMediaContext mediaContext, ref bool changed) {
        if (rule is ICssStyleRule styleRule) {
            builder.AppendLine(styleRule.CssText);
            return;
        }

        if (rule is ICssMediaRule mediaRule) {
            changed = true;
            if (HtmlComputedStyleEngine.IsApplicableMedia(mediaRule.ConditionText, mediaContext)) {
                foreach (ICssRule childRule in mediaRule.Rules) {
                    AppendActiveMediaStyleRule(builder, childRule, mediaContext, ref changed);
                }
            }

            return;
        }

        if (rule is ICssSupportsRule supportsRule) {
            changed = true;
            if (HtmlComputedStyleEngine.IsApplicableSupports(supportsRule.ConditionText)) {
                foreach (ICssRule childRule in supportsRule.Rules) {
                    AppendActiveMediaStyleRule(builder, childRule, mediaContext, ref changed);
                }
            }

            return;
        }

        builder.AppendLine(rule.CssText);
    }
}
