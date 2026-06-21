using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using System.Text.RegularExpressions;

namespace OfficeIMO.Html;

/// <summary>
/// Produces stable, policy-aware normalized HTML for OfficeIMO conversion workflows.
/// </summary>
public static class HtmlNormalizer {
    private const int MaxSrcDocDepth = 8;
    private static readonly Regex CssUrlExpression = new Regex("url\\(\\s*(?:\"(?<url>[^\"]*)\"|'(?<url>[^']*)'|(?<url>[^)]+))\\s*\\)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
    private static readonly Regex CssImportExpression = new Regex("@import\\s+(?:url\\(\\s*(?:\"(?<url>[^\"]*)\"|'(?<url>[^']*)'|(?<url>[^)]+))\\s*\\)|\"(?<url>[^\"]*)\"|'(?<url>[^']*)')", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
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
        return NormalizeDocument(document, options, 0);
    }

    private static string NormalizeDocument(IHtmlDocument document, HtmlNormalizationOptions options, int srcDocDepth) {
        INode root = options.UseBodyContentsOnly
            ? HtmlDocumentParser.GetConversionRoot(document, useBodyContentsOnly: true)
            : document.DocumentElement;

        var builder = new StringBuilder();
        if (!options.UseBodyContentsOnly && root is IElement documentElement) {
            AppendElement(builder, documentElement, options, srcDocDepth);
        } else {
            foreach (INode child in root.ChildNodes) {
                AppendNode(builder, child, options, srcDocDepth, preserveWhitespace: false);
            }
        }

        return builder.ToString().Trim();
    }

    private static void AppendNode(StringBuilder builder, INode node, HtmlNormalizationOptions options, int srcDocDepth, bool preserveWhitespace = false) {
        if (node.NodeType == NodeType.Text) {
            AppendText(builder, node.TextContent, options, preserveWhitespace);
            return;
        }

        if (node.NodeType == NodeType.Comment) {
            if (options.PreserveComments) {
                builder.Append("<!--").Append(node.TextContent).Append("-->");
            }

            return;
        }

        if (node is IElement element) {
            AppendElement(builder, element, options, srcDocDepth, preserveWhitespace);
        }
    }

    private static void AppendElement(StringBuilder builder, IElement element, HtmlNormalizationOptions options, int srcDocDepth, bool preserveWhitespace = false) {
        string name = element.TagName.ToLowerInvariant();
        if (SkippedElements.Contains(name) || (name == "style" && !options.PreserveStyleElements)) {
            return;
        }

        string emittedName = IsForeignContent(element) ? element.TagName : name;
        builder.Append('<').Append(emittedName);
        foreach (KeyValuePair<string, string> attribute in NormalizeAttributes(element, options, srcDocDepth)) {
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
                    builder.Append(EscapeRawTextElementContent(styleText, "style"));
                }
            } else {
                bool childPreserveWhitespace = preserveWhitespace || IsPreformattedElement(name);
                foreach (INode child in element.ChildNodes) {
                    AppendNode(builder, child, options, srcDocDepth, childPreserveWhitespace);
                }
            }

            builder.Append("</").Append(emittedName).Append('>');
        }
    }

    private static IReadOnlyList<KeyValuePair<string, string>> NormalizeAttributes(IElement element, HtmlNormalizationOptions options, int srcDocDepth) {
        var attributes = new List<KeyValuePair<string, string>>();
        bool preserveAttributeCasing = IsForeignContent(element);
        foreach (IAttr attribute in element.Attributes) {
            string name = attribute.Name.ToLowerInvariant();
            if (options.RemoveEventHandlerAttributes && name.StartsWith("on", StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            string value = NormalizeAttributeValue(element, name, attribute.Value, options, srcDocDepth);
            if (IsUrlAttribute(element, name) || SrcSetAttributes.Contains(name)) {
                if (string.IsNullOrWhiteSpace(value)) {
                    continue;
                }
            }

            string emittedName = preserveAttributeCasing ? attribute.Name : name;
            attributes.Add(new KeyValuePair<string, string>(emittedName, value));
        }

        return attributes
            .OrderBy(pair => AttributeOrder(pair.Key))
            .ThenBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static string NormalizeAttributeValue(IElement element, string name, string value, HtmlNormalizationOptions options, int srcDocDepth) {
        if (BooleanAttributes.Contains(name) && IsBooleanValue(name, value)) {
            return name;
        }

        if (string.Equals(name, "srcdoc", StringComparison.OrdinalIgnoreCase)) {
            return NormalizeSrcDoc(value, options, srcDocDepth);
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

        return value;
    }

    private static string NormalizeSrcDoc(string value, HtmlNormalizationOptions options, int srcDocDepth) {
        if (string.IsNullOrWhiteSpace(value) || srcDocDepth >= MaxSrcDocDepth) {
            return string.Empty;
        }

        IHtmlDocument nested = HtmlDocumentParser.ParseDocument(value);
        HtmlNormalizationOptions nestedOptions = CopyOptions(options);
        nestedOptions.BaseUri = HtmlDocumentParser.ResolveEffectiveBaseUri(nested, options.BaseUri);
        nestedOptions.UseBodyContentsOnly = true;
        return NormalizeDocument(nested, nestedOptions, srcDocDepth + 1);
    }

    private static HtmlNormalizationOptions CopyOptions(HtmlNormalizationOptions options) {
        return new HtmlNormalizationOptions {
            BaseUri = options.BaseUri,
            UrlPolicy = options.UrlPolicy,
            UseBodyContentsOnly = options.UseBodyContentsOnly,
            PreserveComments = options.PreserveComments,
            PreserveStyleElements = options.PreserveStyleElements,
            RemoveEventHandlerAttributes = options.RemoveEventHandlerAttributes,
            CollapseTextWhitespace = options.CollapseTextWhitespace
        };
    }

    private static void AppendText(StringBuilder builder, string? text, HtmlNormalizationOptions options, bool preserveWhitespace) {
        if (string.IsNullOrEmpty(text)) {
            return;
        }

        string value = options.CollapseTextWhitespace && !preserveWhitespace
            ? CollapseWhitespaceRuns(text!)
            : text!;
        if (value.Length > 0) {
            if (options.CollapseTextWhitespace && !preserveWhitespace) {
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

    private static bool IsPreformattedElement(string name) {
        return string.Equals(name, "pre", StringComparison.OrdinalIgnoreCase)
            || string.Equals(name, "textarea", StringComparison.OrdinalIgnoreCase);
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

        var replacements = new List<CssReplacement>();
        foreach (Match match in CssUrlExpression.Matches(css)) {
            if (!IsCssFunctionNameAt(css, match.Index, "url") || IsInsideCssString(css, match.Index)) {
                continue;
            }

            string source = match.Groups["url"].Value.Trim().Trim('\'', '"');
            string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(source, baseUri, policy);
            string replacement = string.IsNullOrWhiteSpace(resolved)
                ? "url(\"\")"
                : "url(\"" + EscapeCssString(resolved) + "\")";
            replacements.Add(new CssReplacement(match.Index, match.Index + match.Length, replacement));
        }

        AddCssStringResourceReplacements(css, CssImportExpression, baseUri, policy, replacements);
        AddImageSetStringResourceReplacements(css, baseUri, policy, replacements);
        return ApplyCssReplacements(css, replacements);
    }

    private static void AddCssStringResourceReplacements(string css, Regex expression, Uri? baseUri, HtmlUrlPolicy policy, ICollection<CssReplacement> replacements) {
        foreach (Match match in expression.Matches(css)) {
            if (IsInsideCssString(css, match.Index)) {
                continue;
            }

            Group group = match.Groups["url"];
            if (!group.Success) {
                continue;
            }

            string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(group.Value.Trim(), baseUri, policy);
            replacements.Add(new CssReplacement(group.Index, group.Index + group.Length, EscapeCssString(resolved)));
        }
    }

    private static void AddImageSetStringResourceReplacements(string css, Uri? baseUri, HtmlUrlPolicy policy, ICollection<CssReplacement> replacements) {
        int index = 0;
        while (index < css.Length) {
            int imageSetIndex = css.IndexOf("image-set", index, StringComparison.OrdinalIgnoreCase);
            if (imageSetIndex < 0) {
                return;
            }

            if (IsInsideCssString(css, imageSetIndex) || !IsImageSetFunction(css, imageSetIndex)) {
                index = imageSetIndex + 9;
                continue;
            }

            int open = css.IndexOf('(', imageSetIndex);
            if (open < 0) {
                return;
            }

            int close = FindMatchingCssParenthesis(css, open);
            if (close <= open) {
                return;
            }

            int cursor = open + 1;
            while (cursor < close) {
                char current = css[cursor];
                if ((current == '"' || current == '\'') && !IsCssTypeFunctionString(css, cursor)) {
                    if (TryReadCssQuotedValue(css, cursor, out string source, out int end)) {
                        string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(source, baseUri, policy);
                        replacements.Add(new CssReplacement(cursor + 1, end - 1, EscapeCssString(resolved)));
                        cursor = end;
                        continue;
                    }
                }

                cursor++;
            }

            index = close + 1;
        }
    }

    private static string ApplyCssReplacements(string css, IEnumerable<CssReplacement> replacements) {
        var ordered = replacements
            .OrderByDescending(range => range.Start)
            .ToList();
        var builder = new StringBuilder(css);
        foreach (CssReplacement replacement in ordered) {
            if (OverlapsLaterReplacement(replacement, ordered)) {
                continue;
            }

            builder.Remove(replacement.Start, replacement.End - replacement.Start);
            builder.Insert(replacement.Start, replacement.Value);
        }

        return builder.ToString();
    }

    private static bool OverlapsLaterReplacement(CssReplacement replacement, IReadOnlyList<CssReplacement> replacements) {
        foreach (CssReplacement other in replacements) {
            if (other == replacement) {
                return false;
            }

            if (other.Start <= replacement.Start && other.End > replacement.Start) {
                return true;
            }
        }

        return false;
    }

    private static string EscapeCssString(string value) {
        return value.Replace("\\", "\\\\").Replace("\"", "\\\"");
    }

    private static bool IsCssFunctionNameAt(string css, int index, string functionName) {
        if (!StartsWith(css, index, functionName)) {
            return false;
        }

        int afterName = index + functionName.Length;
        if (afterName >= css.Length || css[afterName] != '(') {
            return false;
        }

        return index == 0 || !IsCssIdentifierCharacter(css[index - 1]);
    }

    private static bool IsImageSetFunction(string css, int imageSetIndex) {
        const string ImageSet = "image-set";
        const string WebKitImageSet = "-webkit-image-set";
        int functionStart = imageSetIndex;
        int nameLength = ImageSet.Length;
        int prefixedStart = imageSetIndex - (WebKitImageSet.Length - ImageSet.Length);
        if (StartsWith(css, prefixedStart, WebKitImageSet)) {
            functionStart = prefixedStart;
            nameLength = WebKitImageSet.Length;
        }

        if (functionStart > 0 && IsCssIdentifierCharacter(css[functionStart - 1])) {
            return false;
        }

        int afterName = functionStart + nameLength;
        return afterName >= css.Length || !IsCssIdentifierCharacter(css[afterName]);
    }

    private static bool IsCssTypeFunctionString(string css, int quoteIndex) {
        int cursor = quoteIndex - 1;
        while (cursor >= 0 && char.IsWhiteSpace(css[cursor])) {
            cursor--;
        }

        if (cursor < 0 || css[cursor] != '(') {
            return false;
        }

        cursor--;
        while (cursor >= 0 && char.IsWhiteSpace(css[cursor])) {
            cursor--;
        }

        int end = cursor + 1;
        while (cursor >= 0 && (char.IsLetter(css[cursor]) || css[cursor] == '-')) {
            cursor--;
        }

        string functionName = css.Substring(cursor + 1, end - cursor - 1);
        return string.Equals(functionName, "type", StringComparison.OrdinalIgnoreCase);
    }

    private static int FindMatchingCssParenthesis(string css, int open) {
        int depth = 0;
        char quote = '\0';
        for (int i = open; i < css.Length; i++) {
            char current = css[i];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(css, i)) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
                continue;
            }

            if (current == '(') {
                depth++;
                continue;
            }

            if (current == ')') {
                depth--;
                if (depth == 0) {
                    return i;
                }
            }
        }

        return -1;
    }

    private static bool TryReadCssQuotedValue(string css, int cursor, out string value, out int end) {
        char quote = css[cursor];
        int start = cursor + 1;
        cursor = start;
        while (cursor < css.Length) {
            if (css[cursor] == quote && !IsEscaped(css, cursor)) {
                value = css.Substring(start, cursor - start);
                end = cursor + 1;
                return true;
            }

            cursor++;
        }

        value = string.Empty;
        end = cursor;
        return false;
    }

    private static bool StartsWith(string text, int index, string value) {
        return index >= 0
            && index + value.Length <= text.Length
            && string.Compare(text, index, value, 0, value.Length, StringComparison.OrdinalIgnoreCase) == 0;
    }

    private static bool IsCssIdentifierCharacter(char value) {
        return char.IsLetterOrDigit(value)
            || value == '_'
            || value == '-'
            || value >= 0x80;
    }

    private static bool IsInsideCssString(string css, int index) {
        char quote = '\0';
        for (int i = 0; i < index && i < css.Length; i++) {
            char current = css[i];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(css, i)) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
            }
        }

        return quote != '\0';
    }

    private static bool IsEscaped(string text, int index) {
        int slashCount = 0;
        for (int i = index - 1; i >= 0 && text[i] == '\\'; i--) {
            slashCount++;
        }

        return slashCount % 2 == 1;
    }

    private static string EscapeRawTextElementContent(string value, string elementName) {
        return Regex.Replace(
            value,
            "</\\s*" + Regex.Escape(elementName),
            match => "<\\/" + match.Value.Substring(2),
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    }

    private static bool IsForeignContent(IElement element) {
        IElement? current = element;
        while (current != null) {
            if (string.Equals(current.TagName, "svg", StringComparison.OrdinalIgnoreCase)
                || string.Equals(current.TagName, "math", StringComparison.OrdinalIgnoreCase)
                || string.Equals(current.NamespaceUri, "http://www.w3.org/2000/svg", StringComparison.Ordinal)
                || string.Equals(current.NamespaceUri, "http://www.w3.org/1998/Math/MathML", StringComparison.Ordinal)) {
                return true;
            }

            current = current.ParentElement;
        }

        return false;
    }

    private static bool IsBooleanValue(string name, string? value) {
        return string.IsNullOrEmpty(value)
            || string.Equals(value, name, StringComparison.OrdinalIgnoreCase);
    }

    private static int AttributeOrder(string name) {
        switch (name.ToLowerInvariant()) {
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

    private sealed class CssReplacement {
        internal CssReplacement(int start, int end, string value) {
            Start = start;
            End = end;
            Value = value;
        }

        internal int Start { get; }
        internal int End { get; }
        internal string Value { get; }
    }
}
