using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using System.Text.RegularExpressions;

namespace OfficeIMO.Html;

/// <summary>
/// Produces stable, policy-aware normalized HTML for OfficeIMO conversion workflows.
/// </summary>
public static class HtmlNormalizer {
    private static readonly Regex CssUrlExpression = new Regex("url\\(\\s*(?:\"(?<url>[^\"]*)\"|'(?<url>[^']*)'|(?<url>[^)]+))\\s*\\)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
    private static readonly HashSet<string> BooleanAttributes = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "allowfullscreen", "async", "autofocus", "autoplay", "checked", "controls", "default", "defer", "disabled",
        "formnovalidate", "hidden", "loop", "multiple", "muted", "nomodule", "novalidate", "open", "readonly",
        "required", "reversed", "selected"
    };
    private static readonly HashSet<string> VoidElements = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "source", "track", "wbr"
    };
    private static readonly HashSet<string> SkippedElements = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "script", "template", "noscript"
    };
    private static readonly HashSet<string> UrlAttributes = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "action", "background", "cite", "data-poster", "data-src", "formaction", "href", "poster", "src", "xlink:href"
    };
    private static readonly HashSet<string> SrcSetAttributes = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "data-srcset", "imagesrcset", "srcset"
    };
    private static readonly char[] WhitespaceSeparators = { ' ', '\t', '\r', '\n', '\f' };

    /// <summary>
    /// Parses and normalizes raw HTML.
    /// </summary>
    public static string Normalize(string html, HtmlNormalizationOptions? options = null) {
        if (html == null) {
            throw new ArgumentNullException(nameof(html));
        }

        return Normalize(HtmlDocumentParser.ParseDocument(html), options);
    }

    /// <summary>
    /// Normalizes an already parsed HTML document.
    /// </summary>
    public static string Normalize(IHtmlDocument document, HtmlNormalizationOptions? options = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        options ??= new HtmlNormalizationOptions();
        INode root = options.UseBodyContentsOnly
            ? HtmlDocumentParser.GetConversionRoot(document, useBodyContentsOnly: true)
            : document.DocumentElement;

        var builder = new StringBuilder();
        if (!options.UseBodyContentsOnly && root is IElement documentElement) {
            AppendElement(builder, documentElement, options);
        } else {
            foreach (INode child in root.ChildNodes) {
                AppendNode(builder, child, options);
            }
        }

        return builder.ToString().Trim();
    }

    private static void AppendNode(StringBuilder builder, INode node, HtmlNormalizationOptions options) {
        if (node.NodeType == NodeType.Text) {
            AppendText(builder, node.TextContent, options);
            return;
        }

        if (node.NodeType == NodeType.Comment) {
            if (options.PreserveComments) {
                builder.Append("<!--").Append(node.TextContent).Append("-->");
            }

            return;
        }

        if (node is IElement element) {
            AppendElement(builder, element, options);
        }
    }

    private static void AppendElement(StringBuilder builder, IElement element, HtmlNormalizationOptions options) {
        string name = element.TagName.ToLowerInvariant();
        if (SkippedElements.Contains(name) || (name == "style" && !options.PreserveStyleElements)) {
            return;
        }

        builder.Append('<').Append(name);
        foreach (KeyValuePair<string, string> attribute in NormalizeAttributes(element, options)) {
            builder.Append(' ').Append(attribute.Key);
            if (!BooleanAttributes.Contains(attribute.Key) || !IsBooleanValue(attribute.Key, attribute.Value)) {
                builder.Append("=\"").Append(WebUtility.HtmlEncode(attribute.Value)).Append('"');
            }
        }

        builder.Append('>');
        if (!VoidElements.Contains(name)) {
            if (name == "style" && options.PreserveStyleElements) {
                string styleText = NormalizeCssUrls(element.TextContent, options.BaseUri, options.UrlPolicy);
                if (styleText.Length > 0) {
                    builder.Append(WebUtility.HtmlEncode(styleText));
                }
            } else {
                foreach (INode child in element.ChildNodes) {
                    AppendNode(builder, child, options);
                }
            }

            builder.Append("</").Append(name).Append('>');
        }
    }

    private static IReadOnlyList<KeyValuePair<string, string>> NormalizeAttributes(IElement element, HtmlNormalizationOptions options) {
        var attributes = new List<KeyValuePair<string, string>>();
        foreach (IAttr attribute in element.Attributes) {
            string name = attribute.Name.ToLowerInvariant();
            if (options.RemoveEventHandlerAttributes && name.StartsWith("on", StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            string value = NormalizeAttributeValue(element, name, attribute.Value, options);
            if (IsUrlAttribute(element, name) || SrcSetAttributes.Contains(name)) {
                if (string.IsNullOrWhiteSpace(value)) {
                    continue;
                }
            }

            attributes.Add(new KeyValuePair<string, string>(name, value));
        }

        return attributes
            .OrderBy(pair => AttributeOrder(pair.Key))
            .ThenBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static string NormalizeAttributeValue(IElement element, string name, string value, HtmlNormalizationOptions options) {
        if (BooleanAttributes.Contains(name) && IsBooleanValue(name, value)) {
            return name;
        }

        if (SrcSetAttributes.Contains(name)) {
            return HtmlImageSourceResolver.ResolveNormalizedSrcSet(value, options.BaseUri, options.UrlPolicy);
        }

        if (IsUrlAttribute(element, name)) {
            return HtmlUrlPolicyEvaluator.ResolveUrl(value, options.BaseUri, options.UrlPolicy);
        }

        if (string.Equals(name, "style", StringComparison.OrdinalIgnoreCase)) {
            return NormalizeCssUrls(value, options.BaseUri, options.UrlPolicy).Trim();
        }

        if (string.Equals(name, "class", StringComparison.OrdinalIgnoreCase)) {
            return string.Join(" ", value.Split(WhitespaceSeparators, StringSplitOptions.RemoveEmptyEntries));
        }

        return value.Trim();
    }

    private static void AppendText(StringBuilder builder, string? text, HtmlNormalizationOptions options) {
        if (string.IsNullOrEmpty(text)) {
            return;
        }

        string value = options.CollapseTextWhitespace
            ? CollapseWhitespaceRuns(text!)
            : text!;
        if (value.Length > 0) {
            if (options.CollapseTextWhitespace) {
                if (value == " ") {
                    if (builder.Length == 0 || char.IsWhiteSpace(builder[builder.Length - 1])) {
                        return;
                    }
                } else if (value[0] == ' ' && (builder.Length == 0 || char.IsWhiteSpace(builder[builder.Length - 1]))) {
                    value = value.TrimStart();
                }
            }

            builder.Append(WebUtility.HtmlEncode(value));
        }
    }

    private static bool IsUrlAttribute(IElement element, string name) {
        if (UrlAttributes.Contains(name)) {
            return true;
        }

        return string.Equals(name, "data", StringComparison.OrdinalIgnoreCase)
            && string.Equals(element.TagName, "object", StringComparison.OrdinalIgnoreCase);
    }

    private static string CollapseWhitespaceRuns(string text) {
        var builder = new StringBuilder(text.Length);
        bool inWhitespace = false;
        for (int i = 0; i < text.Length; i++) {
            char current = text[i];
            if (char.IsWhiteSpace(current)) {
                if (!inWhitespace) {
                    builder.Append(' ');
                    inWhitespace = true;
                }

                continue;
            }

            builder.Append(current);
            inWhitespace = false;
        }

        return builder.ToString();
    }

    private static string NormalizeCssUrls(string css, Uri? baseUri, HtmlUrlPolicy policy) {
        if (string.IsNullOrWhiteSpace(css)) {
            return string.Empty;
        }

        return CssUrlExpression.Replace(css, match => {
            string source = match.Groups["url"].Value.Trim().Trim('\'', '"');
            string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(source, baseUri, policy);
            return string.IsNullOrWhiteSpace(resolved)
                ? "url(\"\")"
                : "url(\"" + EscapeCssString(resolved) + "\")";
        });
    }

    private static string EscapeCssString(string value) {
        return value.Replace("\\", "\\\\").Replace("\"", "\\\"");
    }

    private static bool IsBooleanValue(string name, string? value) {
        return string.IsNullOrEmpty(value)
            || string.Equals(value, name, StringComparison.OrdinalIgnoreCase);
    }

    private static int AttributeOrder(string name) {
        switch (name) {
            case "id":
                return 0;
            case "class":
                return 1;
            case "href":
            case "src":
            case "srcset":
                return 2;
            case "alt":
            case "title":
            case "aria-label":
                return 3;
            case "style":
                return 9;
            default:
                return 5;
        }
    }
}
