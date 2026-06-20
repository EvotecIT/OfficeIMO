using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using System.Text.RegularExpressions;

namespace OfficeIMO.Html;

/// <summary>
/// Shared resource discovery and policy planning for OfficeIMO HTML workflows.
/// </summary>
public static class HtmlResourcePipeline {
    private const int MaxSrcDocDepth = 8;
    private const string ResourceSelector = "image, meta[http-equiv], [src], [srcset], [href], [xlink\\:href], [data], [data-src], [data-srcset], [poster], [data-poster], [action], [formaction], [background], [srcdoc], [imagesrcset]";
    private static readonly Regex CssUrlExpression = new Regex("url\\(\\s*(?:\"(?<url>[^\"]+)\"|'(?<url>[^']+)'|(?<url>[^)]+))\\s*\\)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
    private static readonly Regex CssVarExpression = new Regex("var\\(\\s*(?<name>--[A-Za-z0-9_-]+)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

    /// <summary>
    /// Parses raw HTML and builds a resource manifest.
    /// </summary>
    public static HtmlResourceManifest BuildManifest(string html, HtmlResourcePipelineOptions? options = null) {
        IHtmlDocument document = HtmlDocumentParser.ParseDocument(html);
        return BuildManifest(document, options);
    }

    /// <summary>
    /// Builds a resource manifest from a parsed document.
    /// </summary>
    public static HtmlResourceManifest BuildManifest(IHtmlDocument document, HtmlResourcePipelineOptions? options = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        options = options ?? new HtmlResourcePipelineOptions();
        Uri? baseUri = HtmlDocumentParser.ResolveEffectiveBaseUri(document, options.BaseUri);
        var manifest = new HtmlResourceManifest();
        foreach (IElement element in document.QuerySelectorAll(ResourceSelector)) {
            AddElementResources(manifest, element, baseUri, options, 0);
        }

        AddCssResources(manifest, document, baseUri, options);
        return manifest;
    }

    private static void AddElementResources(HtmlResourceManifest manifest, IElement element, Uri? baseUri, HtmlResourcePipelineOptions options, int srcDocDepth) {
        string name = element.TagName.ToLowerInvariant();
        AddLegacyBackground(manifest, element, baseUri, options);
        switch (name) {
            case "img":
                AddImage(manifest, element, baseUri, options);
                break;
            case "image":
                AddAttribute(manifest, HtmlResourceKind.Image, element, "href", baseUri, options);
                AddAttribute(manifest, HtmlResourceKind.Image, element, "xlink:href", baseUri, options);
                AddAttribute(manifest, HtmlResourceKind.Image, element, "src", baseUri, options);
                break;
            case "source":
                AddSource(manifest, element, baseUri, options);
                break;
            case "link":
                AddLink(manifest, element, baseUri, options);
                break;
            case "base":
                break;
            case "meta":
                AddMetaRefresh(manifest, element, baseUri, options);
                break;
            case "a":
            case "area":
                AddAttribute(manifest, HtmlResourceKind.Hyperlink, element, "href", baseUri, options);
                break;
            case "form":
                AddAttribute(manifest, HtmlResourceKind.Hyperlink, element, "action", baseUri, options);
                break;
            case "input":
                AddSubmitterFormAction(manifest, element, baseUri, options);
                if (string.Equals(element.GetAttribute("type"), "image", StringComparison.OrdinalIgnoreCase)) {
                    AddImage(manifest, element, baseUri, options);
                }

                break;
            case "button":
                AddSubmitterFormAction(manifest, element, baseUri, options);
                break;
            case "script":
                if (IsExecutableScriptElement(element)) {
                    AddAttribute(manifest, HtmlResourceKind.Script, element, "src", baseUri, options);
                }

                break;
            case "video":
                AddAttribute(manifest, HtmlResourceKind.Image, element, "poster", baseUri, options);
                AddAttribute(manifest, HtmlResourceKind.Image, element, "data-poster", baseUri, options);
                AddAttribute(manifest, HtmlResourceKind.Media, element, "src", baseUri, options);
                AddAttribute(manifest, HtmlResourceKind.Media, element, "data-src", baseUri, options);
                break;
            case "audio":
            case "track":
                AddAttribute(manifest, HtmlResourceKind.Media, element, "src", baseUri, options);
                AddAttribute(manifest, HtmlResourceKind.Media, element, "data-src", baseUri, options);
                break;
            case "object":
                AddAttribute(manifest, HtmlResourceKind.Other, element, "data", baseUri, options);
                break;
            case "embed":
                AddAttribute(manifest, HtmlResourceKind.Other, element, "src", baseUri, options);
                break;
            case "iframe":
                if (!element.HasAttribute("srcdoc")) {
                    AddAttribute(manifest, HtmlResourceKind.Other, element, "src", baseUri, options);
                }

                AddSrcDocResources(manifest, element, baseUri, options, srcDocDepth);
                break;
            default:
                AddAttribute(manifest, HtmlResourceKind.Other, element, "src", baseUri, options);
                AddAttribute(manifest, HtmlResourceKind.Other, element, "href", baseUri, options);
                break;
        }
    }

    private static void AddLegacyBackground(HtmlResourceManifest manifest, IElement element, Uri? baseUri, HtmlResourcePipelineOptions options) {
        AddAttribute(manifest, HtmlResourceKind.Image, element, "background", baseUri, options);
    }

    private static void AddSubmitterFormAction(HtmlResourceManifest manifest, IElement element, Uri? baseUri, HtmlResourcePipelineOptions options) {
        if (!element.HasAttribute("formaction")) {
            return;
        }

        string name = element.TagName.ToLowerInvariant();
        string type = (element.GetAttribute("type") ?? string.Empty).Trim();
        bool isSubmitter = string.Equals(name, "button", StringComparison.OrdinalIgnoreCase)
            ? !string.Equals(type, "button", StringComparison.OrdinalIgnoreCase) && !string.Equals(type, "reset", StringComparison.OrdinalIgnoreCase)
            : string.Equals(type, "submit", StringComparison.OrdinalIgnoreCase) || string.Equals(type, "image", StringComparison.OrdinalIgnoreCase);
        if (isSubmitter) {
            AddAttribute(manifest, HtmlResourceKind.Hyperlink, element, "formaction", baseUri, options);
        }
    }

    private static void AddImage(HtmlResourceManifest manifest, IElement element, Uri? baseUri, HtmlResourcePipelineOptions options) {
        foreach (string attribute in new[] { "data-src", "src" }) {
            AddAttribute(manifest, HtmlResourceKind.Image, element, attribute, baseUri, options);
        }

        AddSrcSet(manifest, HtmlResourceKind.Image, element, "srcset", baseUri, options);
        AddSrcSet(manifest, HtmlResourceKind.Image, element, "data-srcset", baseUri, options);
    }

    private static void AddSource(HtmlResourceManifest manifest, IElement element, Uri? baseUri, HtmlResourcePipelineOptions options) {
        string parentName = element.ParentElement?.TagName.ToLowerInvariant() ?? string.Empty;
        switch (parentName) {
            case "picture":
                AddSrcSet(manifest, HtmlResourceKind.Image, element, "srcset", baseUri, options);
                AddSrcSet(manifest, HtmlResourceKind.Image, element, "data-srcset", baseUri, options);
                break;
            case "audio":
            case "video":
                AddAttribute(manifest, HtmlResourceKind.Media, element, "src", baseUri, options);
                AddAttribute(manifest, HtmlResourceKind.Media, element, "data-src", baseUri, options);
                break;
        }
    }

    private static void AddLink(HtmlResourceManifest manifest, IElement element, Uri? baseUri, HtmlResourcePipelineOptions options) {
        string rel = element.GetAttribute("rel") ?? string.Empty;
        HashSet<string> relTokens = GetRelTokens(rel);
        HtmlResourceKind kind;
        if (relTokens.Contains("stylesheet")) {
            kind = HtmlResourceKind.Stylesheet;
        } else if (relTokens.Contains("modulepreload")) {
            kind = HtmlResourceKind.Script;
        } else if (relTokens.Contains("preload")) {
            kind = GetPreloadKind(element.GetAttribute("as"));
        } else if (relTokens.Contains("font")) {
            kind = HtmlResourceKind.Font;
        } else if (relTokens.Contains("icon") || relTokens.Contains("apple-touch-icon") || relTokens.Contains("shortcut icon")) {
            kind = HtmlResourceKind.Image;
        } else {
            kind = HtmlResourceKind.Hyperlink;
        }

        AddAttribute(manifest, kind, element, "href", baseUri, options);
        if (relTokens.Contains("preload") && kind == HtmlResourceKind.Image) {
            AddSrcSet(manifest, HtmlResourceKind.Image, element, "imagesrcset", baseUri, options);
        }
    }

    private static HashSet<string> GetRelTokens(string rel) {
        var tokens = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string token in rel.Split(new[] { ' ', '\t', '\r', '\n', '\f' }, StringSplitOptions.RemoveEmptyEntries)) {
            string normalized = token.Trim();
            if (normalized.Length > 0) {
                tokens.Add(normalized);
            }
        }

        if (tokens.Contains("shortcut") && tokens.Contains("icon")) {
            tokens.Add("shortcut icon");
        }

        return tokens;
    }

    private static HtmlResourceKind GetPreloadKind(string? asAttribute) {
        switch ((asAttribute ?? string.Empty).Trim().ToLowerInvariant()) {
            case "script":
            case "worker":
            case "serviceworker":
                return HtmlResourceKind.Script;
            case "style":
                return HtmlResourceKind.Stylesheet;
            case "image":
                return HtmlResourceKind.Image;
            case "font":
                return HtmlResourceKind.Font;
            case "audio":
            case "track":
            case "video":
                return HtmlResourceKind.Media;
            default:
                return HtmlResourceKind.Other;
        }
    }

    private static void AddCssResources(HtmlResourceManifest manifest, IHtmlDocument document, Uri? baseUri, HtmlResourcePipelineOptions options) {
        foreach (IElement styleElement in document.QuerySelectorAll("style")) {
            if (!IsCssStyleElement(styleElement)) {
                continue;
            }

            AddCssReferences(manifest, styleElement, "css", styleElement.TextContent, baseUri, options);
        }

        foreach (IElement element in document.QuerySelectorAll("[style]")) {
            AddCssReferences(manifest, element, "style", element.GetAttribute("style") ?? string.Empty, baseUri, options);
        }
    }

    private static bool IsCssStyleElement(IElement styleElement) {
        string type = (styleElement.GetAttribute("type") ?? string.Empty).Trim();
        if (type.Length == 0) {
            return true;
        }

        int parameterStart = type.IndexOf(';');
        if (parameterStart >= 0) {
            type = type.Substring(0, parameterStart).Trim();
        }

        return string.Equals(type, "text/css", StringComparison.OrdinalIgnoreCase);
    }

    private static void AddSrcDocResources(HtmlResourceManifest manifest, IElement element, Uri? baseUri, HtmlResourcePipelineOptions options, int srcDocDepth) {
        string? srcdoc = element.GetAttribute("srcdoc");
        if (string.IsNullOrWhiteSpace(srcdoc)) {
            return;
        }

        if (srcDocDepth >= MaxSrcDocDepth) {
            return;
        }

        IHtmlDocument nested = HtmlDocumentParser.ParseDocument(srcdoc!);
        Uri? nestedBaseUri = HtmlDocumentParser.ResolveEffectiveBaseUri(nested, baseUri);
        foreach (IElement nestedElement in nested.QuerySelectorAll(ResourceSelector)) {
            AddElementResources(manifest, nestedElement, nestedBaseUri, options, srcDocDepth + 1);
        }

        AddCssResources(manifest, nested, nestedBaseUri, options);
    }

    private static void AddCssReferences(HtmlResourceManifest manifest, IElement element, string attributeName, string css, Uri? baseUri, HtmlResourcePipelineOptions options) {
        if (string.IsNullOrWhiteSpace(css)) {
            return;
        }

        css = StripCssCommentsOutsideStrings(css);
        bool scanImports = !string.Equals(attributeName, "style", StringComparison.OrdinalIgnoreCase);
        Dictionary<string, List<CssCustomPropertyUrl>> customPropertyUrls = ExtractCustomPropertyUrls(css);
        List<SourceRange> resolvedVarFallbackRanges = GetResolvedVarFallbackRanges(css, customPropertyUrls);
        var importRanges = new List<SourceRange>();
        if (scanImports) {
            foreach (CssImportReference reference in ExtractCssImports(css)) {
                string source = reference.Source;
                if (!string.IsNullOrWhiteSpace(source)) {
                    importRanges.Add(new SourceRange(reference.Start, reference.End));
                    AddRaw(manifest, HtmlResourceKind.Stylesheet, element, attributeName + "-import", DecodeCssEscapes(source), baseUri, options);
                }
            }
        }

        AddUsedCustomPropertyUrls(manifest, element, attributeName, css, customPropertyUrls, baseUri, options);
        foreach (CssStringUrlReference reference in ExtractImageSetStringUrls(css)) {
            if (!TryGetCustomPropertyName(css, reference.Start, out _) && IsSupportedCssUrlDeclaration(css, reference.Start)) {
                AddRaw(manifest, ClassifyCssUrl(css, reference.Start), element, attributeName + "-image-set", DecodeCssEscapes(reference.Source), baseUri, options);
            }
        }

        foreach (Match match in CssUrlExpression.Matches(css)) {
            string source = match.Groups["url"].Value.Trim().Trim('\'', '"');
            if (!string.IsNullOrWhiteSpace(source)
                && IsCssFunctionNameAt(css, match.Index, "url")
                && !IsImportUrl(match.Index, importRanges)
                && !IsImportUrl(match.Index, resolvedVarFallbackRanges)
                && !IsImportAtRuleUrl(css, match.Index)
                && !IsAtRulePreludeUrl(css, match.Index)
                && !IsInsideCssString(css, match.Index)
                && !IsCustomPropertyUrl(css, match.Index)
                && IsSupportedCssUrlDeclaration(css, match.Index)) {
                AddRaw(manifest, ClassifyCssUrl(css, match.Index), element, attributeName + "-url", DecodeCssEscapes(source), baseUri, options);
            }
        }
    }

    private static Dictionary<string, List<CssCustomPropertyUrl>> ExtractCustomPropertyUrls(string css) {
        var urls = new Dictionary<string, List<CssCustomPropertyUrl>>(StringComparer.Ordinal);
        foreach (Match match in CssUrlExpression.Matches(css)) {
            if (TryGetCustomPropertyName(css, match.Index, out string propertyName)) {
                AddCustomPropertyUrl(urls, propertyName, DecodeCssEscapes(match.Groups["url"].Value.Trim().Trim('\'', '"')), GetDeclarationSelector(css, match.Index), GetDeclarationStart(css, match.Index));
            }
        }

        foreach (CssStringUrlReference reference in ExtractImageSetStringUrls(css)) {
            if (TryGetCustomPropertyName(css, reference.Start, out string propertyName)) {
                AddCustomPropertyUrl(urls, propertyName, DecodeCssEscapes(reference.Source), GetDeclarationSelector(css, reference.Start), GetDeclarationStart(css, reference.Start));
            }
        }

        return urls;
    }

    private static void AddCustomPropertyUrl(IDictionary<string, List<CssCustomPropertyUrl>> urls, string propertyName, string source, string selector, int declarationStart) {
        if (string.IsNullOrWhiteSpace(source)) {
            return;
        }

        if (!urls.TryGetValue(propertyName, out List<CssCustomPropertyUrl>? values)) {
            values = new List<CssCustomPropertyUrl>();
            urls[propertyName] = values;
        }

        values.Add(new CssCustomPropertyUrl(source, selector, declarationStart));
    }

    private static void AddUsedCustomPropertyUrls(HtmlResourceManifest manifest, IElement element, string attributeName, string css, IReadOnlyDictionary<string, List<CssCustomPropertyUrl>> customPropertyUrls, Uri? baseUri, HtmlResourcePipelineOptions options) {
        if (customPropertyUrls.Count == 0) {
            return;
        }

        foreach (Match match in CssVarExpression.Matches(css)) {
            string propertyName = match.Groups["name"].Value;
            if (!IsCssFunctionNameAt(css, match.Index, "var") || IsInsideCssString(css, match.Index) || !customPropertyUrls.TryGetValue(propertyName, out List<CssCustomPropertyUrl>? sources)) {
                continue;
            }

            HtmlResourceKind kind = ClassifyCssUrl(css, match.Index);
            if (kind == HtmlResourceKind.Other) {
                continue;
            }

            string useSelector = GetDeclarationSelector(css, match.Index);
            int selectedDeclarationStart = -1;
            foreach (CssCustomPropertyUrl source in sources) {
                if (CanSubstituteCustomProperty(source.Selector, useSelector) && source.DeclarationStart >= selectedDeclarationStart) {
                    selectedDeclarationStart = source.DeclarationStart;
                }
            }

            if (selectedDeclarationStart >= 0) {
                foreach (CssCustomPropertyUrl source in sources) {
                    if (source.DeclarationStart == selectedDeclarationStart && CanSubstituteCustomProperty(source.Selector, useSelector)) {
                        AddRaw(manifest, kind, element, attributeName + "-var-url", source.Source, baseUri, options);
                    }
                }
            }
        }
    }

    private static bool CanSubstituteCustomProperty(string definitionSelector, string useSelector) {
        if (string.IsNullOrWhiteSpace(definitionSelector)) {
            return string.IsNullOrWhiteSpace(useSelector);
        }

        if (string.Equals(definitionSelector, useSelector, StringComparison.OrdinalIgnoreCase)) {
            return true;
        }

        foreach (string definitionPart in definitionSelector.Split(',')) {
            string normalizedDefinition = definitionPart.Trim();
            if (string.Equals(normalizedDefinition, ":root", StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalizedDefinition, "html", StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalizedDefinition, "body", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            foreach (string usePart in useSelector.Split(',')) {
                string normalizedUse = usePart.Trim();
                if (IsAncestorSelector(normalizedDefinition, normalizedUse) || IsSameElementSelectorPrefix(normalizedDefinition, normalizedUse)) {
                    return true;
                }
            }
        }

        return false;
    }

    private static List<SourceRange> GetResolvedVarFallbackRanges(string css, IReadOnlyDictionary<string, List<CssCustomPropertyUrl>> customPropertyUrls) {
        var ranges = new List<SourceRange>();
        if (customPropertyUrls.Count == 0) {
            return ranges;
        }

        foreach (Match match in CssVarExpression.Matches(css)) {
            string propertyName = match.Groups["name"].Value;
            if (!IsCssFunctionNameAt(css, match.Index, "var") || IsInsideCssString(css, match.Index) || !customPropertyUrls.TryGetValue(propertyName, out List<CssCustomPropertyUrl>? sources)) {
                continue;
            }

            int open = css.IndexOf('(', match.Index);
            if (open < 0) {
                continue;
            }

            int close = FindMatchingCssParenthesis(css, open);
            if (close <= open) {
                continue;
            }

            string useSelector = GetDeclarationSelector(css, match.Index);
            if (!sources.Any(source => CanSubstituteCustomProperty(source.Selector, useSelector))) {
                continue;
            }

            int comma = FindTopLevelComma(css, open + 1, close);
            if (comma >= 0) {
                ranges.Add(new SourceRange(comma + 1, close));
            }
        }

        return ranges;
    }

    private static int FindTopLevelComma(string css, int start, int end) {
        int depth = 0;
        char quote = '\0';
        for (int i = start; i < end; i++) {
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
                if (depth > 0) {
                    depth--;
                }

                continue;
            }

            if (depth == 0 && current == ',') {
                return i;
            }
        }

        return -1;
    }

    private static bool IsAncestorSelector(string definitionSelector, string useSelector) {
        if (definitionSelector.Length == 0 || useSelector.Length <= definitionSelector.Length) {
            return false;
        }

        if (!useSelector.StartsWith(definitionSelector, StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        char next = useSelector[definitionSelector.Length];
        return char.IsWhiteSpace(next) || next == '>';
    }

    private static bool IsSameElementSelectorPrefix(string definitionSelector, string useSelector) {
        if (definitionSelector.Length == 0 || useSelector.Length <= definitionSelector.Length) {
            return false;
        }

        if (!useSelector.StartsWith(definitionSelector, StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        char next = useSelector[definitionSelector.Length];
        return next == '.'
            || next == '#'
            || next == '['
            || next == ':';
    }

    private static string GetDeclarationSelector(string css, int index) {
        int blockStart = css.LastIndexOf('{', Math.Max(0, index - 1));
        if (blockStart < 0) {
            return string.Empty;
        }

        int previousBlockEnd = css.LastIndexOf('}', Math.Max(0, blockStart - 1));
        int previousStatementEnd = css.LastIndexOf(';', Math.Max(0, blockStart - 1));
        int selectorStart = Math.Max(0, Math.Max(previousBlockEnd, previousStatementEnd) + 1);
        string selector = css.Substring(selectorStart, blockStart - selectorStart).Trim();
        int groupingStart = selector.LastIndexOf('{');
        return groupingStart >= 0
            ? selector.Substring(groupingStart + 1).Trim()
            : selector;
    }

    private static int GetDeclarationStart(string css, int index) {
        int blockStart = css.LastIndexOf('{', Math.Max(0, index - 1));
        int previousStatementEnd = css.LastIndexOf(';', Math.Max(0, index - 1));
        return Math.Max(0, Math.Max(blockStart, previousStatementEnd) + 1);
    }

    private static IEnumerable<CssStringUrlReference> ExtractImageSetStringUrls(string css) {
        int index = 0;
        while (index < css.Length) {
            int functionStart = css.IndexOf("image-set", index, StringComparison.OrdinalIgnoreCase);
            if (functionStart < 0) {
                yield break;
            }

            if (IsInsideCssString(css, functionStart)) {
                index = functionStart + 9;
                continue;
            }

            if (!TryGetImageSetFunction(css, functionStart, out int nameStart, out int nameLength)) {
                index = functionStart + 9;
                continue;
            }

            int cursor = SkipWhitespace(css, nameStart + nameLength);
            if (cursor >= css.Length || css[cursor] != '(') {
                index = functionStart + 9;
                continue;
            }

            int close = FindMatchingCssParenthesis(css, cursor);
            if (close <= cursor) {
                yield break;
            }

            int valueCursor = cursor + 1;
            while (valueCursor < close) {
                char current = css[valueCursor];
                if ((current == '"' || current == '\'') && !IsCssTypeFunctionString(css, valueCursor)) {
                    if (TryReadCssQuotedValue(css, valueCursor, out string source, out int end)) {
                        if (!string.IsNullOrWhiteSpace(source)) {
                            yield return new CssStringUrlReference(functionStart, end, source);
                        }

                        valueCursor = end;
                        continue;
                    }
                }

                valueCursor++;
            }

            index = close + 1;
        }
    }

    private static bool TryGetImageSetFunction(string css, int imageSetIndex, out int functionStart, out int nameLength) {
        const string ImageSet = "image-set";
        const string WebKitImageSet = "-webkit-image-set";

        functionStart = imageSetIndex;
        nameLength = ImageSet.Length;

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

    private static IEnumerable<CssImportReference> ExtractCssImports(string css) {
        int index = 0;
        while (index < css.Length) {
            int importStart = css.IndexOf("@import", index, StringComparison.OrdinalIgnoreCase);
            if (importStart < 0) {
                yield break;
            }

            if (IsInsideCssString(css, importStart)) {
                index = importStart + 7;
                continue;
            }

            if (!HasImportTokenBoundary(css, importStart)) {
                index = importStart + 7;
                continue;
            }

            if (HasStyleRuleBefore(css, importStart)) {
                yield break;
            }

            int cursor = SkipWhitespace(css, importStart + 7);
            string source;
            int end;
            if (StartsWith(css, cursor, "url(")) {
                cursor = SkipWhitespace(css, cursor + 4);
                if (!TryReadCssUrlFunctionSource(css, cursor, out source, out end)) {
                    index = importStart + 7;
                    continue;
                }
            } else if (cursor < css.Length && (css[cursor] == '"' || css[cursor] == '\'')) {
                if (!TryReadCssQuotedValue(css, cursor, out source, out end)) {
                    index = importStart + 7;
                    continue;
                }
            } else {
                int sourceStart = cursor;
                while (cursor < css.Length && !char.IsWhiteSpace(css[cursor]) && css[cursor] != ';') {
                    cursor++;
                }

                source = css.Substring(sourceStart, cursor - sourceStart);
                end = cursor;
            }

            int importEnd = end;
            while (importEnd < css.Length && css[importEnd] != ';') {
                importEnd++;
            }

            if (importEnd < css.Length) {
                importEnd++;
            }

            yield return new CssImportReference(importStart, importEnd, source);
            index = importEnd;
        }
    }

    private static bool TryReadCssUrlFunctionSource(string css, int cursor, out string source, out int end) {
        if (cursor < css.Length && (css[cursor] == '"' || css[cursor] == '\'')) {
            if (!TryReadCssQuotedValue(css, cursor, out source, out cursor)) {
                end = cursor;
                return false;
            }
        } else {
            int sourceStart = cursor;
            while (cursor < css.Length && css[cursor] != ')') {
                cursor++;
            }

            source = css.Substring(sourceStart, cursor - sourceStart).Trim();
        }

        cursor = SkipWhitespace(css, cursor);
        if (cursor < css.Length && css[cursor] == ')') {
            cursor++;
        }

        end = cursor;
        return true;
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

    private static int SkipWhitespace(string text, int index) {
        while (index < text.Length && char.IsWhiteSpace(text[index])) {
            index++;
        }

        return index;
    }

    private static bool StartsWith(string text, int index, string value) {
        return index >= 0
            && index + value.Length <= text.Length
            && string.Compare(text, index, value, 0, value.Length, StringComparison.OrdinalIgnoreCase) == 0;
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

    private static string StripCssCommentsOutsideStrings(string css) {
        var result = new System.Text.StringBuilder(css.Length);
        char quote = '\0';
        for (int i = 0; i < css.Length; i++) {
            char current = css[i];
            if (quote != '\0') {
                result.Append(current);
                if (current == quote && !IsEscaped(css, i)) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
                result.Append(current);
                continue;
            }

            if (current == '/' && i + 1 < css.Length && css[i + 1] == '*') {
                i += 2;
                while (i + 1 < css.Length && !(css[i] == '*' && css[i + 1] == '/')) {
                    i++;
                }

                if (i + 1 < css.Length) {
                    i++;
                }

                result.Append(' ');
                continue;
            }

            result.Append(current);
        }

        return result.ToString();
    }

    private static bool IsCustomPropertyUrl(string css, int index) {
        return TryGetCustomPropertyName(css, index, out _);
    }

    private static bool TryGetCustomPropertyName(string css, int index, out string propertyName) {
        int blockStart = css.LastIndexOf('{', Math.Max(0, index - 1));
        int previousBoundary = Math.Max(css.LastIndexOf(';', Math.Max(0, index - 1)), blockStart);
        string declaration = css.Substring(Math.Max(0, previousBoundary + 1), index - Math.Max(0, previousBoundary + 1)).TrimStart();
        if (!declaration.StartsWith("--", StringComparison.Ordinal)) {
            propertyName = string.Empty;
            return false;
        }

        int separator = declaration.IndexOf(':');
        if (separator <= 0) {
            propertyName = string.Empty;
            return false;
        }

        propertyName = declaration.Substring(0, separator).Trim();
        return propertyName.Length > 2;
    }

    private static bool IsImportAtRuleUrl(string css, int index) {
        int previousSemicolon = css.LastIndexOf(';', Math.Max(0, index - 1));
        int previousBlockEnd = css.LastIndexOf('}', Math.Max(0, index - 1));
        int previousBoundary = Math.Max(previousSemicolon, previousBlockEnd);
        string statement = css.Substring(Math.Max(0, previousBoundary + 1), index - Math.Max(0, previousBoundary + 1));
        int importStart = statement.IndexOf("@import", StringComparison.OrdinalIgnoreCase);
        return importStart >= 0 && HasImportTokenBoundary(statement, importStart);
    }

    private static bool IsAtRulePreludeUrl(string css, int index) {
        int previousOpen = css.LastIndexOf('{', Math.Max(0, index - 1));
        int previousClose = css.LastIndexOf('}', Math.Max(0, index - 1));
        int previousSemicolon = css.LastIndexOf(';', Math.Max(0, index - 1));
        int previousBoundary = Math.Max(previousOpen, Math.Max(previousClose, previousSemicolon));
        int segmentStart = Math.Max(0, previousBoundary + 1);
        string prefix = css.Substring(segmentStart, index - segmentStart);
        if (prefix.LastIndexOf('@') < 0) {
            return false;
        }

        int nextOpen = css.IndexOf('{', index);
        if (nextOpen < 0) {
            return false;
        }

        int nextSemicolon = css.IndexOf(';', index);
        int nextClose = css.IndexOf('}', index);
        return (nextSemicolon < 0 || nextOpen < nextSemicolon)
            && (nextClose < 0 || nextOpen < nextClose);
    }

    private static bool HasImportTokenBoundary(string css, int importStart) {
        int afterImport = importStart + 7;
        return afterImport >= css.Length || char.IsWhiteSpace(css[afterImport]);
    }

    private static bool HasStyleRuleBefore(string css, int index) {
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
                continue;
            }

            if (current == '{' || current == '}') {
                return true;
            }
        }

        return false;
    }

    private static HtmlResourceKind ClassifyCssUrl(string css, int index) {
        string propertyName = GetCssDeclarationPropertyName(css, index);
        int blockStart = css.LastIndexOf('{', Math.Max(0, index - 1));
        string blockPrefix = blockStart >= 0 ? css.Substring(0, blockStart).ToLowerInvariant() : string.Empty;
        int fontFaceStart = blockPrefix.LastIndexOf("@font-face", StringComparison.Ordinal);
        int previousBlockEnd = blockPrefix.LastIndexOf('}');
        if (fontFaceStart >= 0 && fontFaceStart > previousBlockEnd) {
            return HtmlResourceKind.Font;
        }

        if (IsSupportedCssImageUrlProperty(propertyName)) {
            return HtmlResourceKind.Image;
        }

        return HtmlResourceKind.Other;
    }

    private static bool IsSupportedCssUrlDeclaration(string css, int index) {
        return ClassifyCssUrl(css, index) != HtmlResourceKind.Other;
    }

    private static string GetCssDeclarationPropertyName(string css, int index) {
        int declarationStart = GetDeclarationStart(css, index);
        int separator = css.IndexOf(':', declarationStart, Math.Max(0, index - declarationStart));
        if (separator <= declarationStart) {
            return string.Empty;
        }

        return css.Substring(declarationStart, separator - declarationStart).Trim().ToLowerInvariant();
    }

    private static bool IsSupportedCssImageUrlProperty(string propertyName) {
        switch (propertyName) {
            case "background":
            case "background-image":
            case "border-image":
            case "border-image-source":
            case "content":
            case "cursor":
            case "list-style":
            case "list-style-image":
            case "mask":
            case "mask-image":
            case "-webkit-mask":
            case "-webkit-mask-image":
                return true;
            default:
                return false;
        }
    }

    private static bool IsImportUrl(int index, IEnumerable<SourceRange> ranges) {
        foreach (SourceRange range in ranges) {
            if (index >= range.Start && index < range.End) {
                return true;
            }
        }

        return false;
    }

    private static string NormalizeSource(string source) {
        return source.Trim().Trim('\'', '"');
    }

    private static string DecodeCssEscapes(string source) {
        if (source.IndexOf('\\') < 0) {
            return source;
        }

        var result = new System.Text.StringBuilder(source.Length);
        for (int i = 0; i < source.Length; i++) {
            char current = source[i];
            if (current != '\\' || i + 1 >= source.Length) {
                result.Append(current);
                continue;
            }

            int cursor = i + 1;
            int hexStart = cursor;
            while (cursor < source.Length && cursor - hexStart < 6 && IsHexDigit(source[cursor])) {
                cursor++;
            }

            if (cursor > hexStart) {
                string hex = source.Substring(hexStart, cursor - hexStart);
                if (!int.TryParse(hex, System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture, out int codePoint)
                    || codePoint == 0
                    || codePoint > 0x10FFFF
                    || (codePoint >= 0xD800 && codePoint <= 0xDFFF)) {
                    result.Append('\uFFFD');
                } else {
                    result.Append(char.ConvertFromUtf32(codePoint));
                }

                if (cursor < source.Length && char.IsWhiteSpace(source[cursor])) {
                    cursor++;
                }

                i = cursor - 1;
                continue;
            }

            result.Append(source[cursor]);
            i = cursor;
        }

        return result.ToString();
    }

    private static bool IsExecutableScriptElement(IElement element) {
        string type = (element.GetAttribute("type") ?? string.Empty).Trim();
        if (type.Length == 0) {
            return true;
        }

        int parameterStart = type.IndexOf(';');
        if (parameterStart >= 0) {
            type = type.Substring(0, parameterStart).Trim();
        }

        return string.Equals(type, "module", StringComparison.OrdinalIgnoreCase)
            || string.Equals(type, "text/javascript", StringComparison.OrdinalIgnoreCase)
            || string.Equals(type, "application/javascript", StringComparison.OrdinalIgnoreCase)
            || string.Equals(type, "application/ecmascript", StringComparison.OrdinalIgnoreCase)
            || string.Equals(type, "text/ecmascript", StringComparison.OrdinalIgnoreCase)
            || string.Equals(type, "text/jscript", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsHexDigit(char value) {
        return (value >= '0' && value <= '9')
            || (value >= 'a' && value <= 'f')
            || (value >= 'A' && value <= 'F');
    }

    private static bool IsEscaped(string text, int index) {
        int slashCount = 0;
        for (int i = index - 1; i >= 0 && text[i] == '\\'; i--) {
            slashCount++;
        }

        return slashCount % 2 == 1;
    }

    private static void AddSrcSet(HtmlResourceManifest manifest, HtmlResourceKind kind, IElement element, string attributeName, Uri? baseUri, HtmlResourcePipelineOptions options) {
        string? raw = element.GetAttribute(attributeName);
        if (string.IsNullOrWhiteSpace(raw)) {
            return;
        }

        foreach (HtmlSrcSetCandidate candidate in HtmlSrcSetParser.Parse(raw, options.MaxResponsiveImageCandidates)) {
            AddRaw(manifest, kind, element, attributeName, candidate.Url, baseUri, options);
        }
    }

    private static void AddAttribute(HtmlResourceManifest manifest, HtmlResourceKind kind, IElement element, string attributeName, Uri? baseUri, HtmlResourcePipelineOptions options) {
        string? source = element.GetAttribute(attributeName);
        if (!string.IsNullOrWhiteSpace(source)) {
            AddRaw(manifest, kind, element, attributeName, source!, baseUri, options);
        }
    }

    private static void AddMetaRefresh(HtmlResourceManifest manifest, IElement element, Uri? baseUri, HtmlResourcePipelineOptions options) {
        if (!string.Equals(element.GetAttribute("http-equiv"), "refresh", StringComparison.OrdinalIgnoreCase)) {
            return;
        }

        string? content = element.GetAttribute("content");
        if (string.IsNullOrWhiteSpace(content)) {
            return;
        }

        if (!TryReadMetaRefreshUrl(content!, out string source)) {
            return;
        }

        AddRaw(manifest, HtmlResourceKind.Hyperlink, element, "content", source, baseUri, options);
    }

    private static bool TryReadMetaRefreshUrl(string content, out string source) {
        source = string.Empty;
        string[] parts = content.Split(';');
        for (int i = 1; i < parts.Length; i++) {
            string parameter = parts[i].Trim();
            int separator = parameter.IndexOf('=');
            if (separator <= 0) {
                continue;
            }

            string name = parameter.Substring(0, separator).Trim();
            if (!string.Equals(name, "url", StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            source = parameter.Substring(separator + 1).Trim();
            break;
        }

        if (source.Length == 0) {
            return false;
        }

        if (source.Length > 1 && ((source[0] == '"' && source[source.Length - 1] == '"') || (source[0] == '\'' && source[source.Length - 1] == '\''))) {
            source = source.Substring(1, source.Length - 2).Trim();
        }

        return source.Length > 0;
    }

    private static void AddRaw(HtmlResourceManifest manifest, HtmlResourceKind kind, IElement element, string attributeName, string source, Uri? baseUri, HtmlResourcePipelineOptions options) {
        string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(source, baseUri, options.UrlPolicy);
        bool isAllowed = !string.IsNullOrWhiteSpace(resolved) && IsResourceKindSchemeAllowed(kind, resolved);
        manifest.Add(new HtmlResourceReference(
            kind,
            element.TagName.ToLowerInvariant(),
            attributeName,
            source.Trim(),
            resolved,
            isAllowed,
            isAllowed ? string.Empty : GetDiagnosticCode(kind)));
    }

    private static string GetDiagnosticCode(HtmlResourceKind kind) {
        switch (kind) {
            case HtmlResourceKind.Image:
                return "ImageResourceRejectedByPolicy";
            case HtmlResourceKind.Stylesheet:
                return "StylesheetResourceRejectedByPolicy";
            case HtmlResourceKind.Hyperlink:
                return "HyperlinkRejectedByPolicy";
            case HtmlResourceKind.Script:
                return "ScriptResourceRejectedByPolicy";
            case HtmlResourceKind.Media:
                return "MediaResourceRejectedByPolicy";
            case HtmlResourceKind.Font:
                return "FontResourceRejectedByPolicy";
            default:
                return "HtmlResourceRejectedByPolicy";
        }
    }

    private static bool IsResourceKindSchemeAllowed(HtmlResourceKind kind, string resolved) {
        if (kind == HtmlResourceKind.Hyperlink) {
            return true;
        }

        return !Uri.TryCreate(resolved, UriKind.Absolute, out var uri)
            || !uri.Scheme.Equals(Uri.UriSchemeMailto, StringComparison.OrdinalIgnoreCase);
    }

    private sealed class CssImportReference {
        internal CssImportReference(int start, int end, string source) {
            Start = start;
            End = end;
            Source = source;
        }

        internal int Start { get; }
        internal int End { get; }
        internal string Source { get; }
    }

    private sealed class SourceRange {
        internal SourceRange(int start, int end) {
            Start = start;
            End = end;
        }

        internal int Start { get; }
        internal int End { get; }
    }

    private sealed class CssStringUrlReference {
        internal CssStringUrlReference(int start, int end, string source) {
            Start = start;
            End = end;
            Source = source;
        }

        internal int Start { get; }
        internal int End { get; }
        internal string Source { get; }
    }

    private sealed class CssCustomPropertyUrl {
        internal CssCustomPropertyUrl(string source, string selector, int declarationStart) {
            Source = source;
            Selector = selector;
            DeclarationStart = declarationStart;
        }

        internal string Source { get; }
        internal string Selector { get; }
        internal int DeclarationStart { get; }
    }
}
