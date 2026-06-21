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
            IHtmlDocument parsed = HtmlDocumentParser.ParseDocument(html);
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

            return changed && parsed.DocumentElement != null
                ? parsed.DocumentElement.OuterHtml
                : html;
        } catch {
            return html;
        }
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
