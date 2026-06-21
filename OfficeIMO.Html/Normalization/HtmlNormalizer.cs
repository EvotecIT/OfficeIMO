using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using System.Globalization;
using System.Text.RegularExpressions;

namespace OfficeIMO.Html;

/// <summary>
/// Produces stable, policy-aware normalized HTML for OfficeIMO conversion workflows.
/// </summary>
public static class HtmlNormalizer {
    private const int MaxSrcDocDepth = 8;
    private static readonly Regex CssUrlExpression = new Regex("(?<name>(?:[uU]|\\\\0{0,4}(?:75|55)\\s?|\\\\[uU])(?:[rR]|\\\\0{0,4}(?:72|52)\\s?|\\\\[rR])(?:[lL]|\\\\0{0,4}(?:6[cC]|4[cC])\\s?|\\\\[lL]))\\(\\s*(?:\"(?<url>[^\"]*)\"|'(?<url>[^']*)'|(?<url>[^)]+))\\s*\\)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
    private static readonly HashSet<string> BooleanAttributes = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "allowfullscreen", "async", "autofocus", "autoplay", "checked", "controls", "default", "defer", "disabled",
        "formnovalidate", "hidden", "loop", "multiple", "muted", "nomodule", "novalidate", "open", "readonly",
        "required", "reversed", "selected"
    };
    private static readonly HashSet<string> VoidElements = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "source", "track", "wbr"
    };
    private static readonly HashSet<string> SkippedElements = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "script", "template"
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
                string styleText = NormalizeCssUrls(element.TextContent, options.BaseUri, GetResourceUrlPolicy(options.UrlPolicy));
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
                if (string.IsNullOrWhiteSpace(value) && !ShouldPreserveEmptyUrlAttribute(element, name, attribute.Value)) {
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
            return HtmlImageSourceResolver.ResolveNormalizedSrcSet(value, options.BaseUri, GetResourceUrlPolicy(options.UrlPolicy));
        }

        if (IsUrlAttribute(element, name)) {
            HtmlUrlPolicy attributePolicy = GetAttributeUrlPolicy(element, name, options.UrlPolicy);
            if (string.IsNullOrWhiteSpace(value) && ShouldPreserveEmptyUrlAttribute(element, name, value)) {
                return HtmlUrlPolicyEvaluator.ResolveUrl(options.BaseUri?.AbsoluteUri, null, attributePolicy);
            }

            Uri? baseUri = string.Equals(element.TagName, "base", StringComparison.OrdinalIgnoreCase)
                && string.Equals(name, "href", StringComparison.OrdinalIgnoreCase)
                && options.BaseElementBaseUri != null
                    ? options.BaseElementBaseUri
                    : options.BaseUri;
            return HtmlUrlPolicyEvaluator.ResolveUrl(value, baseUri, attributePolicy);
        }

        if (string.Equals(name, "style", StringComparison.OrdinalIgnoreCase)) {
            return NormalizeCssUrls(value, options.BaseUri, GetResourceUrlPolicy(options.UrlPolicy)).Trim();
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
            BaseElementBaseUri = options.BaseElementBaseUri,
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

    private static bool ShouldPreserveEmptyUrlAttribute(IElement element, string name, string value) {
        if (!string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        string tagName = element.TagName.ToLowerInvariant();
        return (string.Equals(name, "href", StringComparison.OrdinalIgnoreCase)
                && (string.Equals(tagName, "a", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(tagName, "area", StringComparison.OrdinalIgnoreCase)))
            || (string.Equals(name, "action", StringComparison.OrdinalIgnoreCase)
                && string.Equals(tagName, "form", StringComparison.OrdinalIgnoreCase))
            || (string.Equals(name, "formaction", StringComparison.OrdinalIgnoreCase)
                && (string.Equals(tagName, "button", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(tagName, "input", StringComparison.OrdinalIgnoreCase)));
    }

    private static HtmlUrlPolicy GetAttributeUrlPolicy(IElement element, string name, HtmlUrlPolicy policy) {
        return IsHyperlinkUrlAttribute(element, name)
            ? policy
            : GetResourceUrlPolicy(policy);
    }

    private static HtmlUrlPolicy GetResourceUrlPolicy(HtmlUrlPolicy policy) {
        HtmlUrlPolicy resourcePolicy = policy.Clone();
        resourcePolicy.AllowMailtoUrls = false;
        return resourcePolicy;
    }

    private static bool IsHyperlinkUrlAttribute(IElement element, string name) {
        string tagName = element.TagName.ToLowerInvariant();
        if (string.Equals(name, "href", StringComparison.OrdinalIgnoreCase)) {
            return string.Equals(tagName, "a", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "area", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "base", StringComparison.OrdinalIgnoreCase);
        }

        if (string.Equals(name, "cite", StringComparison.OrdinalIgnoreCase)) {
            return true;
        }

        if (string.Equals(name, "action", StringComparison.OrdinalIgnoreCase)) {
            return string.Equals(tagName, "form", StringComparison.OrdinalIgnoreCase);
        }

        return string.Equals(name, "formaction", StringComparison.OrdinalIgnoreCase)
            && (string.Equals(tagName, "button", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "input", StringComparison.OrdinalIgnoreCase));
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

            string source = DecodeCssEscapes(match.Groups["url"].Value.Trim().Trim('\'', '"'));
            string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(source, baseUri, policy);
            string replacement = string.IsNullOrWhiteSpace(resolved)
                ? "url(\"\")"
                : "url(\"" + EscapeCssString(resolved) + "\")";
            replacements.Add(new CssReplacement(match.Index, match.Index + match.Length, replacement));
        }

        AddCssImportResourceReplacements(css, baseUri, policy, replacements);
        AddImageSetStringResourceReplacements(css, baseUri, policy, replacements);
        return ApplyCssReplacements(css, replacements);
    }

    private static void AddCssImportResourceReplacements(string css, Uri? baseUri, HtmlUrlPolicy policy, ICollection<CssReplacement> replacements) {
        int index = 0;
        while (index < css.Length) {
            int importStart = css.IndexOf("@import", index, StringComparison.OrdinalIgnoreCase);
            if (importStart < 0) {
                return;
            }

            if (IsInsideCssString(css, importStart) || !HasAtRuleTokenBoundary(css, importStart, "@import")) {
                index = importStart + 7;
                continue;
            }

            int cursor = SkipCssWhitespaceAndComments(css, importStart + 7);
            if (!TryReadCssImportValue(css, cursor, out int sourceStart, out int sourceEnd)) {
                index = importStart + 7;
                continue;
            }

            string source = DecodeCssEscapes(css.Substring(sourceStart, sourceEnd - sourceStart).Trim());
            string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(source, baseUri, policy);
            replacements.Add(new CssReplacement(sourceStart, sourceEnd, EscapeCssString(resolved)));
            index = sourceEnd;
        }
    }

    private static bool TryReadCssImportValue(string css, int cursor, out int sourceStart, out int sourceEnd) {
        sourceStart = 0;
        sourceEnd = 0;
        if (StartsWith(css, cursor, "url(")) {
            cursor = SkipCssWhitespaceAndComments(css, cursor + 4);
            if (cursor < css.Length && (css[cursor] == '"' || css[cursor] == '\'')) {
                if (!TryReadCssQuotedValue(css, cursor, out _, out int end)) {
                    return false;
                }

                sourceStart = cursor + 1;
                sourceEnd = end - 1;
                return true;
            }

            sourceStart = cursor;
            while (cursor < css.Length && css[cursor] != ')') {
                cursor++;
            }

            sourceEnd = TrimCssValueEnd(css, sourceStart, cursor);
            return sourceEnd >= sourceStart;
        }

        if (cursor < css.Length && (css[cursor] == '"' || css[cursor] == '\'')) {
            if (!TryReadCssQuotedValue(css, cursor, out _, out int end)) {
                return false;
            }

            sourceStart = cursor + 1;
            sourceEnd = end - 1;
            return true;
        }

        sourceStart = cursor;
        while (cursor < css.Length && !char.IsWhiteSpace(css[cursor]) && css[cursor] != ';') {
            cursor++;
        }

        sourceEnd = cursor;
        return sourceEnd > sourceStart;
    }

    private static int TrimCssValueEnd(string css, int start, int end) {
        while (end > start && char.IsWhiteSpace(css[end - 1])) {
            end--;
        }

        return end;
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
                        string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(DecodeCssEscapes(source), baseUri, policy);
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
            .ThenByDescending(range => range.End - range.Start)
            .ToList();
        var builder = new StringBuilder(css);
        var applied = new List<CssReplacement>();
        foreach (CssReplacement replacement in ordered) {
            if (applied.Any(appliedReplacement => RangesOverlap(replacement, appliedReplacement))) {
                continue;
            }

            builder.Remove(replacement.Start, replacement.End - replacement.Start);
            builder.Insert(replacement.Start, replacement.Value);
            applied.Add(replacement);
        }

        return builder.ToString();
    }

    private static bool RangesOverlap(CssReplacement first, CssReplacement second) {
        return first.Start < second.End && second.Start < first.End;
    }

    private static string EscapeCssString(string value) {
        return value.Replace("\\", "\\\\").Replace("\"", "\\\"");
    }

    private static bool IsCssFunctionNameAt(string css, int index, string functionName) {
        int open = css.IndexOf('(', index);
        if (open <= index) {
            return false;
        }

        string rawName = css.Substring(index, open - index).Trim();
        if (!string.Equals(DecodeCssEscapes(rawName), functionName, StringComparison.OrdinalIgnoreCase)) {
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

    private static bool HasAtRuleTokenBoundary(string css, int index, string token) {
        int after = index + token.Length;
        return (index == 0 || !IsCssIdentifierCharacter(css[index - 1]))
            && (after >= css.Length || !IsCssIdentifierCharacter(css[after]));
    }

    private static int SkipCssWhitespaceAndComments(string css, int index) {
        while (index < css.Length) {
            if (char.IsWhiteSpace(css[index])) {
                index++;
                continue;
            }

            if (index + 1 < css.Length && css[index] == '/' && css[index + 1] == '*') {
                int commentEnd = css.IndexOf("*/", index + 2, StringComparison.Ordinal);
                if (commentEnd < 0) {
                    return css.Length;
                }

                index = commentEnd + 2;
                continue;
            }

            break;
        }

        return index;
    }

    private static bool IsCssIdentifierCharacter(char value) {
        return char.IsLetterOrDigit(value)
            || value == '_'
            || value == '-'
            || value >= 0x80;
    }

    private static string DecodeCssEscapes(string source) {
        if (source.IndexOf('\\') < 0) {
            return source;
        }

        var result = new StringBuilder(source.Length);
        for (int i = 0; i < source.Length; i++) {
            char current = source[i];
            if (current != '\\' || i + 1 >= source.Length) {
                result.Append(current);
                continue;
            }

            int cursor = i + 1;
            if (!IsHexDigit(source[cursor])) {
                result.Append(source[cursor]);
                i = cursor;
                continue;
            }

            int hexStart = cursor;
            while (cursor < source.Length && cursor - hexStart < 6 && IsHexDigit(source[cursor])) {
                cursor++;
            }

            string hex = source.Substring(hexStart, cursor - hexStart);
            if (int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int codePoint)
                && codePoint > 0
                && codePoint <= 0x10FFFF
                && (codePoint < 0xD800 || codePoint > 0xDFFF)) {
                result.Append(char.ConvertFromUtf32(codePoint));
            }

            if (cursor < source.Length && char.IsWhiteSpace(source[cursor])) {
                cursor++;
            }

            i = cursor - 1;
        }

        return result.ToString();
    }

    private static bool IsHexDigit(char value) {
        return (value >= '0' && value <= '9')
            || (value >= 'a' && value <= 'f')
            || (value >= 'A' && value <= 'F');
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
