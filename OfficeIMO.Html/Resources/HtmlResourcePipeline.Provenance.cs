using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Html;

public static partial class HtmlResourcePipeline {
    private const char CssCommentMask = '\u0001';
    internal static bool IsActiveProvenanceStyleElement(IElement element) =>
        IsCssStyleElement(element) && IsPotentiallyApplicableProvenanceMedia(element.GetAttribute("media") ?? string.Empty);

    internal static bool IsApplicableProvenanceMedia(IElement element) =>
        IsPotentiallyApplicableProvenanceMedia(element.GetAttribute("media") ?? string.Empty);

    private static bool IsPotentiallyApplicableProvenanceMedia(string mediaText) {
        var options = new HtmlResourcePipelineOptions();
        return IsApplicableMedia(mediaText, options) ||
            HtmlComputedStyleEngine.IsPotentiallyApplicableScreenMedia(mediaText, options.MediaFeatures);
    }

    internal static bool IsActivePictureImageSource(IElement element) {
        if (!string.Equals(element.ParentElement?.LocalName, "picture", StringComparison.OrdinalIgnoreCase)) return true;
        return AppearsBeforePictureFallback(element) &&
            HasPictureSourceCandidate(element) &&
            IsPotentiallyApplicableProvenanceMedia(element.GetAttribute("media") ?? string.Empty) &&
            IsSupportedPictureSourceType(element.GetAttribute("type"));
    }

    internal static bool IsActivePictureFallbackImage(IElement element) =>
        !HasUnconditionalPictureSourceBeforeFallback(element);

    private static bool AppearsBeforePictureFallback(IElement element) {
        IElement? parent = element.ParentElement;
        if (parent == null) return true;
        foreach (IElement sibling in parent.Children) {
            if (ReferenceEquals(sibling, element)) return true;
            if (string.Equals(sibling.LocalName, "img", StringComparison.OrdinalIgnoreCase)) return false;
        }
        return false;
    }

    private static bool HasUnconditionalPictureSourceBeforeFallback(IElement element) {
        IElement? parent = element.ParentElement;
        if (parent == null || !string.Equals(parent.LocalName, "picture", StringComparison.OrdinalIgnoreCase)) return false;
        foreach (IElement sibling in parent.Children) {
            if (ReferenceEquals(sibling, element)) return false;
            if (!string.Equals(sibling.LocalName, "source", StringComparison.OrdinalIgnoreCase) ||
                !HasPictureSourceCandidate(sibling) ||
                !IsSupportedPictureSourceType(sibling.GetAttribute("type"))) continue;
            string media = (sibling.GetAttribute("media") ?? string.Empty).Trim(' ', '\t', '\n', '\f', '\r');
            if (media.Length == 0 || media.Equals("all", StringComparison.OrdinalIgnoreCase) ||
                media.Equals("screen", StringComparison.OrdinalIgnoreCase)) return true;
        }
        return false;
    }

    internal static HtmlProvenanceCssScope CollectProvenanceCssImageScope(
        IHtmlDocument document,
        long maximumStylesheetBytes,
        long maximumExpandedBytes) {
        var options = new HtmlResourcePipelineOptions();
        List<HtmlProvenanceDataStylesheet> dataStylesheets = MaterializeDataStylesheets(
            document, maximumStylesheetBytes, maximumExpandedBytes, out long decodedStylesheetBytes);
        try {
            HtmlComputedStyleSet computedStyleSet = HtmlComputedStyleEngine.ComputeForProvenance(document, options);
            IReadOnlyDictionary<IElement, HtmlComputedStyle> computedStyles = computedStyleSet.Elements;
            Dictionary<string, List<CssCustomPropertyDefinition>> documentDefinitions = ExtractDocumentCustomPropertyDefinitions(document, options);
            Dictionary<IElement, int> inlineSourceOrders = GetInlineStyleSourceOrders(document, GetDocumentCssSourceOrder(document));
            var result = new HtmlProvenanceCssScope(computedStyleSet);

            foreach (IElement styleElement in document.QuerySelectorAll("style")) {
                if (!IsCssStyleElement(styleElement) || !IsPotentiallyApplicableProvenanceMedia(styleElement.GetAttribute("media") ?? string.Empty)) continue;
                CollectResolvedCustomPropertyDeclarations(styleElement.TextContent, documentDefinitions, inlineSourceOrders, document, null, computedStyles, options, result.UsedCustomPropertyDeclarations);
                CollectResolvedVarFallbackStarts(styleElement.TextContent, documentDefinitions, inlineSourceOrders, document, null, computedStyles, options, "css", styleElement, result);
            }
            foreach (IElement element in document.QuerySelectorAll("[style]")) {
                string css = element.GetAttribute("style") ?? string.Empty;
                int sourceOrderBase = inlineSourceOrders.TryGetValue(element, out int sourceOrder)
                    ? sourceOrder
                    : GetDocumentCssSourceOrder(document);
                List<SourceRange> inactiveRanges = GetInactiveCssRuleRanges(
                    css,
                    options,
                    includePotentialResponsiveScreenMedia: true,
                    includeProvenanceImageSupports: true);
                Dictionary<string, List<CssCustomPropertyDefinition>> definitions = MergeCustomPropertyDefinitions(
                    documentDefinitions,
                    ExtractInlineCustomPropertyDefinitions(element, inlineSourceOrders, options, includeSelf: false));
                definitions = MergeCustomPropertyDefinitions(definitions,
                    ExtractCustomPropertyDefinitions(css, inactiveRanges, sourceOrderBase, isInline: true, sourceOwner: element));
                CollectResolvedCustomPropertyDeclarations(css, definitions, inlineSourceOrders, document, element, computedStyles, options, result.UsedCustomPropertyDeclarations);
                CollectResolvedVarFallbackStarts(css, definitions, inlineSourceOrders, document, element, computedStyles, options, "style", element, result);
            }
            result.DecodedStylesheetBytes = decodedStylesheetBytes;
            foreach (HtmlProvenanceDataStylesheet stylesheet in dataStylesheets) {
                result.DataStylesheets.Add(stylesheet.Link, stylesheet);
                RemapStylesheetOwner(result.UsedCustomPropertyDeclarations, stylesheet.MaterializedStyle, stylesheet.Link);
                RemapStylesheetOwner(result.ResolvedVarFallbackStarts, stylesheet.MaterializedStyle, stylesheet.Link);
            }
            return result;
        } finally {
            foreach (HtmlProvenanceDataStylesheet stylesheet in dataStylesheets) stylesheet.MaterializedStyle.Remove();
        }
    }

    private static List<HtmlProvenanceDataStylesheet> MaterializeDataStylesheets(
        IHtmlDocument document,
        long maximumStylesheetBytes,
        long maximumExpandedBytes,
        out long decodedStylesheetBytes) {
        var stylesheets = new List<HtmlProvenanceDataStylesheet>();
        long expandedBytes = 0;
        decodedStylesheetBytes = 0;
        try {
            foreach (IElement link in document.QuerySelectorAll("link[href]")) {
                string href = link.GetAttribute("href") ?? string.Empty;
                int commaIndex = href.IndexOf(',');
                int fragmentIndex = commaIndex >= 0 ? href.IndexOf('#', commaIndex + 1) : -1;
                string fragment = fragmentIndex >= 0 ? href.Substring(fragmentIndex) : string.Empty;
                string dataSource = fragmentIndex >= 0 ? href.Substring(0, fragmentIndex) : href;
                if (!IsHtmlStylesheetLink(link) ||
                    !IsPotentiallyApplicableProvenanceMedia(link.GetAttribute("media") ?? string.Empty) ||
                    !HtmlDataUri.TryParse(dataSource, out HtmlDataUri dataUri) ||
                    !string.Equals(dataUri.MediaType, "text/css", StringComparison.OrdinalIgnoreCase)) continue;

                long decodedByteCount;
                string css;
                try {
                    decodedByteCount = dataUri.EstimateDecodedByteCount();
                    if (decodedByteCount > maximumStylesheetBytes) {
                        throw new InvalidDataException("An embedded HTML stylesheet exceeds the configured asset limit.");
                    }
                    expandedBytes = checked(expandedBytes + decodedByteCount);
                    if (expandedBytes > maximumExpandedBytes) {
                        throw new InvalidDataException("Embedded HTML stylesheets exceed the configured expanded-container limit.");
                    }
                    css = dataUri.DecodeText();
                    decodedStylesheetBytes = expandedBytes;
                } catch (OverflowException exception) {
                    throw new InvalidDataException("Embedded HTML stylesheets declare an invalid expanded size.", exception);
                } catch (UriFormatException) {
                    continue;
                } catch (FormatException) {
                    continue;
                } catch (ArgumentException) {
                    continue;
                }
                if (string.IsNullOrWhiteSpace(css) || link.Parent == null) continue;

                IElement style = document.CreateElement("style");
                style.TextContent = css;
                string media = link.GetAttribute("media") ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(media)) style.SetAttribute("media", media);
                INode parent = link.Parent;
                INode? next = link.NextSibling;
                if (next == null) parent.AppendChild(style);
                else parent.InsertBefore(style, next);
                stylesheets.Add(new HtmlProvenanceDataStylesheet(
                    link, style, css, dataUri.Metadata, fragment));
            }
            return stylesheets;
        } catch {
            foreach (HtmlProvenanceDataStylesheet stylesheet in stylesheets) stylesheet.MaterializedStyle.Remove();
            throw;
        }
    }

    private static bool IsHtmlStylesheetLink(IElement link) {
        if (!string.Equals(link.NamespaceUri, "http://www.w3.org/1999/xhtml", StringComparison.Ordinal)) return false;
        bool stylesheet = (link.GetAttribute("rel") ?? string.Empty)
            .Split(new[] { '\t', '\n', '\f', '\r', ' ' }, StringSplitOptions.RemoveEmptyEntries)
            .Any(token => token.Equals("stylesheet", StringComparison.OrdinalIgnoreCase));
        if (!stylesheet) return false;
        string type = link.GetAttribute("type") ?? string.Empty;
        int parameter = type.IndexOf(';');
        if (parameter >= 0) type = type.Substring(0, parameter);
        type = type.Trim(' ', '\t', '\n', '\f', '\r');
        return type.Length == 0 || type.Equals("text/css", StringComparison.OrdinalIgnoreCase);
    }

    private static void RemapStylesheetOwner(
        IDictionary<IElement, HashSet<int>> owners,
        IElement materializedStyle,
        IElement link) {
        if (!owners.TryGetValue(materializedStyle, out HashSet<int>? starts)) return;
        owners.Remove(materializedStyle);
        if (!owners.TryGetValue(link, out HashSet<int>? existing)) owners.Add(link, starts);
        else existing.UnionWith(starts);
    }

    private static void CollectResolvedVarFallbackStarts(
        string css,
        IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> definitions,
        IReadOnlyDictionary<IElement, int> inlineSourceOrders,
        IHtmlDocument document,
        IElement? inlineUseElement,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> computedStyles,
        HtmlResourcePipelineOptions options,
        string attributeName,
        IElement sourceOwner,
        HtmlProvenanceCssScope result) {
        if (string.IsNullOrWhiteSpace(css) || definitions.Count == 0) return;
        string masked = MaskCssComments(css);
        List<SourceRange> inactiveRanges = GetInactiveCssRuleRanges(
            masked,
            options,
            includePotentialResponsiveScreenMedia: true,
            includeProvenanceImageSupports: true);
        foreach (Match match in CssUrlExpression.Matches(masked)) {
            if (!IsValidCssUrlMatch(masked, match)) continue;
            if (IsResolvedVarFallbackUrl(masked, match.Index, definitions, inlineSourceOrders, document,
                inlineUseElement, computedStyles, inactiveRanges, options, attributeName)) {
                result.AddResolvedFallback(sourceOwner, match.Index);
            }
        }
        foreach (CssStringUrlReference reference in ExtractImageSetStringUrls(masked)) {
            if (IsResolvedVarFallbackUrl(masked, reference.Start, definitions, inlineSourceOrders, document,
                inlineUseElement, computedStyles, inactiveRanges, options, attributeName)) {
                result.AddResolvedFallback(sourceOwner, reference.Start);
            }
        }
    }

    private static void CollectResolvedCustomPropertyDeclarations(
        string css,
        IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> definitions,
        IReadOnlyDictionary<IElement, int> inlineSourceOrders,
        IHtmlDocument document,
        IElement? inlineUseElement,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> computedStyles,
        HtmlResourcePipelineOptions options,
        IDictionary<IElement, HashSet<int>> result) {
        if (string.IsNullOrWhiteSpace(css) || definitions.Count == 0) return;
        string masked = MaskCssComments(css);
        List<SourceRange> inactiveRanges = GetInactiveCssRuleRanges(
            masked,
            options,
            includePotentialResponsiveScreenMedia: true,
            includeProvenanceImageSupports: true);
        foreach (Match variable in CssVarExpression.Matches(masked)) {
            if (IsInRanges(variable.Index, inactiveRanges) ||
                !IsCssFunctionNameAt(masked, variable.Index, "var") || IsInsideCssString(masked, variable.Index) ||
                ClassifyCssUrl(masked, variable.Index) != HtmlResourceKind.Image) continue;
            string propertyName = DecodeCssEscapes(variable.Groups["name"].Value);
            string useSelector = GetDeclarationSelector(masked, variable.Index);
            IEnumerable<IElement?> useElements = inlineUseElement != null
                ? new IElement?[] { inlineUseElement }
                : GetElementsMatchingSelectorList(document, useSelector).Cast<IElement?>();
            foreach (IElement? useElement in useElements) {
                IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> effectiveDefinitions = definitions;
                if (inlineUseElement == null && useElement != null) {
                    Dictionary<string, List<CssCustomPropertyDefinition>> inlineDefinitions =
                        ExtractInlineCustomPropertyDefinitions(useElement, inlineSourceOrders, options, includeSelf: true);
                    if (inlineDefinitions.Count != 0) effectiveDefinitions = MergeCustomPropertyDefinitions(definitions, inlineDefinitions);
                }
                foreach (CssCustomPropertyDefinition source in ResolveCustomPropertyUrlDefinitions(
                    propertyName, effectiveDefinitions, useSelector, document, useElement, computedStyles,
                    new HashSet<string>(StringComparer.Ordinal), depth: 0)) {
                    if (source.SourceOwner == null) continue;
                    if (!result.TryGetValue(source.SourceOwner, out HashSet<int>? starts)) {
                        starts = new HashSet<int>();
                        result.Add(source.SourceOwner, starts);
                    }
                    starts.Add(source.LocalDeclarationStart);
                }
            }
        }
    }

    internal static IEnumerable<HtmlCssImageReference> EnumerateProvenanceCssImageReferences(
        IHtmlDocument document,
        string attributeName,
        string css,
        ISet<int>? usedCustomPropertyDeclarationStarts = null,
        ISet<int>? resolvedVarFallbackStarts = null,
        IReadOnlyDictionary<IElement, HtmlComputedStyle>? computedStyles = null,
        HtmlComputedStyleSet? computedStyleSet = null,
        IElement? sourceOwner = null) {
        if (string.IsNullOrWhiteSpace(css)) yield break;
        string masked = MaskCssComments(css);
        List<SourceRange> inactiveRanges = GetInactiveCssRuleRanges(
            masked,
            new HtmlResourcePipelineOptions(),
            includePotentialResponsiveScreenMedia: true,
            includeProvenanceImageSupports: true);
        var emittedRanges = new HashSet<(int Start, int Length)>();
        foreach (Match match in CssUrlExpression.Matches(masked)) {
            if (!IsValidCssUrlMatch(masked, match)) continue;
            bool isCustomProperty = TryGetCustomPropertyName(masked, match.Index, out _);
            if (IsInRanges(match.Index, inactiveRanges) || resolvedVarFallbackStarts?.Contains(match.Index) == true ||
                !IsCssFunctionNameAt(masked, match.Index, "url") ||
                IsInsideCssString(masked, match.Index) ||
                IsImportAtRuleUrl(masked, match.Index) ||
                IsAtRulePreludeUrl(masked, match.Index) ||
                !IsCssReferenceForMatchingSelector(document, attributeName, masked, match.Index, computedStyles) ||
                isCustomProperty && (!TryGetCustomPropertyDeclarationStart(masked, match.Index, out int declarationStart) ||
                    usedCustomPropertyDeclarationStarts == null || !usedCustomPropertyDeclarationStarts.Contains(declarationStart)) ||
                !isCustomProperty && ClassifyCssUrl(masked, match.Index) != HtmlResourceKind.Image) continue;
            Group sourceGroup = match.Groups["url"];
            int leading = 0;
            while (leading < sourceGroup.Length && IsCssWhitespace(sourceGroup.Value[leading])) leading++;
            int trailing = sourceGroup.Length;
            while (trailing > leading && IsCssWhitespace(sourceGroup.Value[trailing - 1])) trailing--;
            if (trailing == leading) continue;
            string source = DecodeCssEscapes(sourceGroup.Value.Substring(leading, trailing - leading));
            if (!isCustomProperty && !IsEffectiveImageDeclaration(
                    document, attributeName, sourceOwner, masked, match.Index, source, computedStyles, computedStyleSet)) continue;
            var range = (sourceGroup.Index + leading, trailing - leading);
            if (emittedRanges.Add(range)) yield return new HtmlCssImageReference(range.Item1, range.Item2, source);
        }

        foreach (CssStringUrlReference reference in ExtractImageSetStringUrls(masked)) {
            bool isCustomProperty = TryGetCustomPropertyName(masked, reference.Start, out _);
            if (IsInRanges(reference.Start, inactiveRanges) || resolvedVarFallbackStarts?.Contains(reference.Start) == true ||
                !IsCssReferenceForMatchingSelector(document, attributeName, masked, reference.Start, computedStyles) ||
                (isCustomProperty
                    ? !TryGetCustomPropertyDeclarationStart(masked, reference.Start, out int declarationStart) ||
                        usedCustomPropertyDeclarationStarts == null || !usedCustomPropertyDeclarationStarts.Contains(declarationStart)
                    : ClassifyCssUrl(masked, reference.Start) != HtmlResourceKind.Image)) continue;
            if (!isCustomProperty && !IsEffectiveImageDeclaration(
                    document, attributeName, sourceOwner, masked, reference.Start,
                    DecodeCssEscapes(reference.Source), computedStyles, computedStyleSet)) continue;
            if (!emittedRanges.Add((reference.SourceStart, reference.Source.Length))) continue;
            yield return new HtmlCssImageReference(reference.SourceStart, reference.Source.Length, DecodeCssEscapes(reference.Source));
        }
    }

    private static bool IsEffectiveImageDeclaration(
        IHtmlDocument document,
        string attributeName,
        IElement? sourceOwner,
        string css,
        int index,
        string source,
        IReadOnlyDictionary<IElement, HtmlComputedStyle>? computedStyles,
        HtmlComputedStyleSet? computedStyleSet) {
        if (TryGetEnclosingKeyframesName(css, index, out _, out _)) return true;
        string propertyName = GetCssDeclarationPropertyName(css, index);
        if (propertyName.Length == 0) return true;
        propertyName = HtmlComputedStyleEngine.GetImageSourcePropertyName(propertyName);
        int declarationStart = GetDeclarationStart(css, index);
        int colon = css.IndexOf(':', declarationStart, Math.Max(0, index - declarationStart));
        if (colon < 0) return true;
        int valueEnd = FindDeclarationValueEnd(css, colon + 1);
        string declarationValue = css.Substring(colon + 1, valueEnd - colon - 1).Trim();
        int important = declarationValue.LastIndexOf("!important", StringComparison.OrdinalIgnoreCase);
        if (important >= 0 && string.IsNullOrWhiteSpace(declarationValue.Substring(important + 10))) {
            declarationValue = declarationValue.Substring(0, important).TrimEnd();
        }
        computedStyles ??= computedStyleSet?.Elements ?? HtmlComputedStyleEngine.Compute(document);
        if (IsInsidePotentialResponsiveMedia(css, index) ||
            sourceOwner != null && IsPotentialResponsiveMediaAttribute(sourceOwner)) return true;
        IEnumerable<IElement> elements;
        string selector = string.Empty;
        if (string.Equals(attributeName, "style", StringComparison.OrdinalIgnoreCase)) {
            if (sourceOwner == null) return true;
            elements = new[] { sourceOwner };
        } else {
            selector = GetDeclarationSelector(css, index);
            elements = GetElementsMatchingSelectorList(document, selector);
        }
        bool sawElement = false;
        foreach (IElement element in elements) {
            sawElement = true;
            if (!computedStyles.TryGetValue(element, out HtmlComputedStyle? elementStyle) ||
                IsDisplayNone(elementStyle)) continue;
            foreach (HtmlComputedStyle style in GetDeclarationStyles(element, elementStyle, selector, computedStyleSet)) {
                if (IsDisplayNone(style)) continue;
                string effective = style.GetValue(propertyName);
                if (effective.Length == 0) {
                    if (!IsInsideContainerRule(css, index)) return true;
                    continue;
                }
                if (ContainsEquivalentImageSource(effective, source) || string.Equals(
                        HtmlRenderCssValues.NormalizeComponentValueWhitespace(DecodeCssEscapes(declarationValue)),
                        HtmlRenderCssValues.NormalizeComponentValueWhitespace(effective),
                        StringComparison.Ordinal)) return true;
            }
        }
        return !sawElement;
    }

    private static bool IsPotentialResponsiveMediaAttribute(IElement element) {
        string mediaText = element.GetAttribute("media") ?? string.Empty;
        var options = new HtmlResourcePipelineOptions();
        return !IsApplicableMedia(mediaText, options) &&
            HtmlComputedStyleEngine.IsPotentiallyApplicableScreenMedia(mediaText, options.MediaFeatures);
    }

    private static bool IsInsidePotentialResponsiveMedia(string css, int sourceIndex) {
        var options = new HtmlResourcePipelineOptions();
        int index = 0;
        while (index < css.Length) {
            int mediaStart = css.IndexOf("@media", index, StringComparison.OrdinalIgnoreCase);
            if (mediaStart < 0) return false;
            if (IsInsideCssString(css, mediaStart) || !HasAtRuleTokenBoundary(css, mediaStart, "@media")) {
                index = mediaStart + 6;
                continue;
            }
            int open = FindNextTopLevelBlockStart(css, mediaStart + 6);
            if (open < 0) return false;
            int close = FindMatchingCssBrace(css, open);
            if (close <= open) return false;
            if (sourceIndex > open && sourceIndex < close) {
                string mediaText = css.Substring(mediaStart + 6, open - mediaStart - 6).Trim();
                if (!IsApplicableMedia(mediaText, options) &&
                    HtmlComputedStyleEngine.IsPotentiallyApplicableScreenMedia(mediaText, options.MediaFeatures)) return true;
                index = open + 1;
                continue;
            }
            index = sourceIndex >= close ? close + 1 : open + 1;
        }
        return false;
    }

    private static IEnumerable<HtmlComputedStyle> GetDeclarationStyles(
        IElement element,
        HtmlComputedStyle elementStyle,
        string selector,
        HtmlComputedStyleSet? computedStyleSet) {
        bool targetsPseudoElement = false;
        bool targetsElement = selector.Length == 0;
        if (computedStyleSet != null && selector.Length != 0) {
            foreach (string selectorPart in SplitTopLevelList(selector)) {
                if (!HtmlComputedStyleEngine.TryParsePseudoElementSelector(
                        selectorPart, out string hostSelector, out HtmlPseudoElementKind kind)) {
                    targetsElement = true;
                    continue;
                }
                targetsPseudoElement = true;
                string normalized = NormalizeSelectorForQuery(hostSelector, stripStatefulPseudoClasses: true);
                if (normalized.Length == 0) normalized = "*";
                bool matches;
                try { matches = element.Matches(normalized); } catch { matches = false; }
                if (!matches || !computedStyleSet.TryGetPseudoStyle(element, kind, out HtmlComputedStyle pseudoStyle) ||
                    !IsGeneratedPseudoElement(pseudoStyle)) continue;
                yield return pseudoStyle;
            }
        }
        if (targetsElement || !targetsPseudoElement) yield return elementStyle;
    }

    private static bool IsDisplayNone(HtmlComputedStyle style) =>
        string.Equals(style.GetValue("display").Trim(), "none", StringComparison.OrdinalIgnoreCase);

    private static bool IsGeneratedPseudoElement(HtmlComputedStyle style) {
        string content = style.GetValue("content").Trim();
        return content.Length != 0 &&
            !content.Equals("normal", StringComparison.OrdinalIgnoreCase) &&
            !content.Equals("none", StringComparison.OrdinalIgnoreCase);
    }

    private static bool ContainsEquivalentImageSource(string effective, string source) {
        foreach (Match match in CssUrlExpression.Matches(effective)) {
            if (!IsValidCssUrlMatch(effective, match)) continue;
            string candidate = DecodeCssEscapes(match.Groups["url"].Value.Trim());
            if (string.Equals(candidate, source, StringComparison.Ordinal)) return true;
        }
        foreach (CssStringUrlReference reference in ExtractImageSetStringUrls(effective)) {
            if (string.Equals(DecodeCssEscapes(reference.Source), source, StringComparison.Ordinal)) return true;
        }
        return false;
    }

    private static bool IsInsideContainerRule(string css, int index) {
        int search = 0;
        while (search < index) {
            int start = css.IndexOf("@container", search, StringComparison.OrdinalIgnoreCase);
            if (start < 0 || start >= index) return false;
            if (IsInsideCssString(css, start) || !HasAtRuleTokenBoundary(css, start, "@container")) {
                search = start + 10;
                continue;
            }
            int open = FindNextTopLevelBlockStart(css, start + 10);
            if (open < 0) return false;
            int close = FindMatchingCssBrace(css, open);
            if (close < 0) return false;
            if (index > open && index < close) return true;
            search = close + 1;
        }
        return false;
    }

    private static bool TryGetCustomPropertyDeclarationStart(string css, int valueIndex, out int declarationStart) {
        declarationStart = -1;
        foreach (Match declaration in CssCustomPropertyDeclarationExpression.Matches(css)) {
            if (declaration.Index > valueIndex) break;
            int colon = css.IndexOf(':', declaration.Index);
            if (colon < 0) continue;
            int end = FindDeclarationValueEnd(css, colon + 1);
            if (valueIndex >= colon + 1 && valueIndex < end) declarationStart = declaration.Index;
        }
        return declarationStart >= 0;
    }

    private static string MaskCssComments(string css) {
        var result = new StringBuilder(css);
        char quote = '\0';
        for (int index = 0; index < css.Length; index++) {
            char current = css[index];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(css, index)) quote = '\0';
                continue;
            }
            if (current is '"' or '\'') { quote = current; continue; }
            if (current != '/' || index + 1 >= css.Length || css[index + 1] != '*') continue;
            result[index++] = CssCommentMask;
            result[index] = CssCommentMask;
            while (index + 1 < css.Length && !(css[index] == '*' && css[index + 1] == '/')) result[++index] = CssCommentMask;
            if (index + 1 < css.Length) { result[index] = CssCommentMask; result[++index] = CssCommentMask; }
        }
        return result.ToString();
    }
}

internal sealed class HtmlProvenanceCssScope {
    internal HtmlProvenanceCssScope(HtmlComputedStyleSet computedStyleSet) {
        ComputedStyleSet = computedStyleSet;
    }

    internal HtmlComputedStyleSet ComputedStyleSet { get; }
    internal IReadOnlyDictionary<IElement, HtmlComputedStyle> ComputedStyles => ComputedStyleSet.Elements;
    internal Dictionary<IElement, HashSet<int>> UsedCustomPropertyDeclarations { get; } = new();
    internal Dictionary<IElement, HashSet<int>> ResolvedVarFallbackStarts { get; } = new();
    internal Dictionary<IElement, HtmlProvenanceDataStylesheet> DataStylesheets { get; } = new();
    internal long DecodedStylesheetBytes { get; set; }

    internal void AddResolvedFallback(IElement owner, int start) {
        if (!ResolvedVarFallbackStarts.TryGetValue(owner, out HashSet<int>? starts)) {
            starts = new HashSet<int>();
            ResolvedVarFallbackStarts.Add(owner, starts);
        }
        starts.Add(start);
    }
}

internal sealed class HtmlProvenanceDataStylesheet {
    internal HtmlProvenanceDataStylesheet(
        IElement link,
        IElement materializedStyle,
        string css,
        string metadata,
        string fragment) {
        Link = link;
        MaterializedStyle = materializedStyle;
        Css = css;
        Metadata = metadata;
        Fragment = fragment;
    }

    internal IElement Link { get; }
    internal IElement MaterializedStyle { get; }
    internal string Css { get; }
    internal string Metadata { get; }
    internal string Fragment { get; }
}

internal readonly struct HtmlCssImageReference {
    internal HtmlCssImageReference(int start, int length, string value) {
        Start = start;
        Length = length;
        Value = value;
    }

    internal int Start { get; }
    internal int Length { get; }
    internal string Value { get; }
}
