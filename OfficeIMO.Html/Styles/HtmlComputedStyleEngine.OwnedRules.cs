namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private static Dictionary<string, Queue<OfficeIMO.Html.Css.HtmlCssQualifiedRule>> IndexOwnedQualifiedRules(
        OfficeIMO.Html.Css.HtmlCssStyleSheet sheet,
        OfficeIMO.Html.Css.HtmlCssNamespaceContext namespaceContext,
        MediaEnvironment environment,
        HtmlCssProcessingBudget budget) {
        var result = new Dictionary<string, Queue<OfficeIMO.Html.Css.HtmlCssQualifiedRule>>(StringComparer.Ordinal);
        IndexOwnedQualifiedRules(sheet.Rules, null, namespaceContext, environment, budget, result, 1);
        return result;
    }

    private static void IndexOwnedQualifiedRules(
        IReadOnlyList<OfficeIMO.Html.Css.HtmlCssRule> rules,
        IReadOnlyList<string>? parentSelectors,
        OfficeIMO.Html.Css.HtmlCssNamespaceContext namespaceContext,
        MediaEnvironment environment,
        HtmlCssProcessingBudget budget,
        IDictionary<string, Queue<OfficeIMO.Html.Css.HtmlCssQualifiedRule>> result,
        int depth) {
        budget.RecordNestingDepth(depth);
        foreach (OfficeIMO.Html.Css.HtmlCssRule rule in rules) {
            if (rule is OfficeIMO.Html.Css.HtmlCssQualifiedRule qualifiedRule) {
                IReadOnlyList<string> resolvedSelectors = ResolveNestedSelectors(
                    RestoreManagedPseudoElements(qualifiedRule.PreludeText), parentSelectors, budget);
                string? selector = OwnedSelectorListIdentity(resolvedSelectors, namespaceContext);
                if (selector != null) {
                    if (!result.TryGetValue(selector, out Queue<OfficeIMO.Html.Css.HtmlCssQualifiedRule>? queue)) {
                        queue = new Queue<OfficeIMO.Html.Css.HtmlCssQualifiedRule>();
                        result.Add(selector, queue);
                    }
                    queue.Enqueue(qualifiedRule);
                }
                if (qualifiedRule.Rules.Count > 0) {
                    IndexOwnedQualifiedRules(qualifiedRule.Rules, resolvedSelectors, namespaceContext,
                        environment, budget, result, depth + 1);
                }
                continue;
            }

            if (rule is not OfficeIMO.Html.Css.HtmlCssAtRule atRule
                || atRule.Rules.Count == 0
                || !ShouldIndexOwnedGroupingRule(atRule, environment)) continue;
            IndexOwnedQualifiedRules(atRule.Rules, parentSelectors, namespaceContext,
                environment, budget, result, depth + 1);
        }
    }

    private static bool ShouldIndexOwnedGroupingRule(
        OfficeIMO.Html.Css.HtmlCssAtRule rule,
        MediaEnvironment environment) {
        if (rule.Name.Equals("media", StringComparison.OrdinalIgnoreCase)) {
            return IsApplicableMedia(rule.PreludeText, environment);
        }
        if (rule.Name.Equals("supports", StringComparison.OrdinalIgnoreCase)) {
            return IsApplicableSupports(rule.PreludeText);
        }
        return rule.Name.Equals("layer", StringComparison.OrdinalIgnoreCase)
            || rule.Name.Equals("container", StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryAddInterleavedOwnedStyleRules(
        AngleSharp.Css.Dom.ICssStyleRule styleRule,
        IReadOnlyList<string> providerResolvedSelectors,
        IReadOnlyList<string>? ownedResolvedSelectors,
        ICollection<StyleRule> rules,
        ParsedSelectorRegistry parsedRuleMatches,
        MediaEnvironment environment,
        HtmlCssProcessingBudget budget,
        CascadeLayerRegistry layers,
        IDictionary<string, Queue<OfficeIMO.Html.Css.HtmlCssQualifiedRule>> ownedRules,
        OfficeIMO.Html.Css.HtmlCssNamespaceContext namespaceContext,
        int depth,
        string? currentLayer,
        IReadOnlyList<ContainerRuleCondition>? containerConditions,
        OfficeIMO.Html.Css.HtmlCssQualifiedRule? ownedRule,
        bool canUseProviderSelectorObjects) {
        if (ownedRule?.Block?.IsClosed != true || styleRule.Rules.Length == 0
            || ownedResolvedSelectors == null) return false;

        var boundaryRules = new List<OfficeIMO.Html.Css.HtmlCssQualifiedRule?>();
        var declarationRuns = new List<List<OfficeIMO.Html.Css.HtmlCssDeclaration>>();
        var currentRun = new List<OfficeIMO.Html.Css.HtmlCssDeclaration>();
        bool sawNestedRule = false;
        bool hasDeclarationAfterNestedRule = false;
        foreach (OfficeIMO.Html.Css.HtmlCssSyntaxNode node in ownedRule.Contents) {
            if (node is OfficeIMO.Html.Css.HtmlCssDeclaration declaration) {
                currentRun.Add(declaration);
                hasDeclarationAfterNestedRule |= sawNestedRule;
                continue;
            }
            declarationRuns.Add(currentRun);
            currentRun = new List<OfficeIMO.Html.Css.HtmlCssDeclaration>();
            boundaryRules.Add(node as OfficeIMO.Html.Css.HtmlCssQualifiedRule);
            sawNestedRule = true;
        }
        declarationRuns.Add(currentRun);
        if (!hasDeclarationAfterNestedRule) return false;

        var ownedNestedRules = boundaryRules.Where(rule => rule != null).Select(rule => rule!).ToArray();
        var providerNestedRules = styleRule.Rules.OfType<AngleSharp.Css.Dom.ICssStyleRule>().ToArray();
        if (ownedNestedRules.Length != providerNestedRules.Length) return false;
        for (int index = 0; index < ownedNestedRules.Length; index++) {
            AngleSharp.Css.Dom.ICssStyleRule providerNestedRule = providerNestedRules[index];
            IReadOnlyList<string> ownedNestedSelectors = ResolveNestedSelectors(
                RestoreManagedPseudoElements(ownedNestedRules[index].PreludeText), ownedResolvedSelectors, budget);
            IReadOnlyList<string> providerNestedSelectors = ResolveNestedSelectors(
                RestoreManagedPseudoElements(providerNestedRule.SelectorText ?? string.Empty), ownedResolvedSelectors, budget);
            string? ownedIdentity = OwnedSelectorListIdentity(ownedNestedSelectors, namespaceContext);
            string? providerIdentity = OwnedSelectorListIdentity(providerNestedSelectors, namespaceContext);
            if (ownedIdentity != null && providerIdentity != null
                && !string.Equals(ownedIdentity, providerIdentity, StringComparison.Ordinal)) return false;
        }

        var parsedDeclarationRuns = new List<Dictionary<string, StyleDeclaration>>(declarationRuns.Count);
        foreach (List<OfficeIMO.Html.Css.HtmlCssDeclaration> run in declarationRuns) {
            Dictionary<string, StyleDeclaration>? parsed = TryCreateOwnedDeclarations(run);
            if (parsed == null) return false;
            parsedDeclarationRuns.Add(parsed);
        }

        bool recordedParentRule = false;
        int nestedRuleIndex = 0;
        for (int index = 0; index < parsedDeclarationRuns.Count; index++) {
            Dictionary<string, StyleDeclaration> declarations = parsedDeclarationRuns[index];
            if (declarations.Count > 0) {
                AddStyleRule(styleRule, providerResolvedSelectors, rules, parsedRuleMatches, budget,
                    currentLayer == null ? null : layers.GetOrder(currentLayer), currentLayer,
                    containerConditions, ownedRule, namespaceContext, canUseProviderSelectorObjects,
                    declarations, recordParsedRule: !recordedParentRule,
                    ownedResolvedSelectors: ownedResolvedSelectors);
                recordedParentRule = true;
            }
            if (index < boundaryRules.Count && boundaryRules[index] is OfficeIMO.Html.Css.HtmlCssQualifiedRule nestedRule) {
                AddStyleRules(providerNestedRules[nestedRuleIndex++], rules, parsedRuleMatches, environment, budget,
                    layers, ownedRules, namespaceContext, depth + 1, currentLayer, ownedResolvedSelectors,
                    containerConditions, nestedRule);
            }
        }
        return true;
    }

    private static OfficeIMO.Html.Css.HtmlCssQualifiedRule? TryTakeOwnedQualifiedRule(
        IDictionary<string, Queue<OfficeIMO.Html.Css.HtmlCssQualifiedRule>> rules,
        IReadOnlyList<string> resolvedSelectors,
        OfficeIMO.Html.Css.HtmlCssNamespaceContext namespaceContext) {
        string? key = OwnedSelectorListIdentity(resolvedSelectors, namespaceContext);
        return key != null && rules.TryGetValue(key, out Queue<OfficeIMO.Html.Css.HtmlCssQualifiedRule>? queue) && queue.Count > 0
            ? queue.Dequeue()
            : null;
    }

    private static void TryConsumeOwnedQualifiedRule(
        IDictionary<string, Queue<OfficeIMO.Html.Css.HtmlCssQualifiedRule>> rules,
        IReadOnlyList<string> resolvedSelectors,
        OfficeIMO.Html.Css.HtmlCssNamespaceContext namespaceContext,
        OfficeIMO.Html.Css.HtmlCssQualifiedRule expected) {
        string? key = OwnedSelectorListIdentity(resolvedSelectors, namespaceContext);
        if (key == null || !rules.TryGetValue(key, out Queue<OfficeIMO.Html.Css.HtmlCssQualifiedRule>? queue)
            || queue.Count == 0 || !ReferenceEquals(queue.Peek(), expected)) return;
        queue.Dequeue();
    }

    private static string? OwnedSelectorListIdentity(
        IReadOnlyList<string> selectors,
        OfficeIMO.Html.Css.HtmlCssNamespaceContext namespaceContext) =>
        OwnedSelectorListIdentity(string.Join(",", selectors), namespaceContext);

    private static string? OwnedSelectorListIdentity(
        string selectorText,
        OfficeIMO.Html.Css.HtmlCssNamespaceContext namespaceContext) {
        var hosts = new List<string>();
        foreach (string selectorTextItem in SplitSelectorList(selectorText)) {
            string selector = selectorTextItem.Trim();
            if (TryParsePseudoElementSelector(selector, out string host, out _)) selector = host;
            if (selector.Length == 0) return null;
            hosts.Add(selector);
        }
        if (hosts.Count == 0) return null;
        try {
            OfficeIMO.Html.Css.HtmlCssSelectorListParseResult parsed = OfficeIMO.Html.Css.HtmlCssSelectorParser.ParseList(
                string.Join(",", hosts), new OfficeIMO.Html.Css.HtmlCssSelectorOptions { Namespaces = namespaceContext });
            return parsed.SelectorList?.CompatibilityIdentity;
        } catch (OfficeIMO.Html.Css.HtmlCssSelectorLimitException) { return null; }
    }

    private static bool CanUseOwnedSelectors(
        IReadOnlyList<string> selectors,
        OfficeIMO.Html.Css.HtmlCssNamespaceContext namespaceContext) {
        if (selectors.Count == 0) return false;
        string source = string.Join(",", selectors.Select(selector =>
            TryParsePseudoElementSelector(selector, out string host, out _) ? host : selector));
        try {
            return OfficeIMO.Html.Css.HtmlCssSelectorParser.ParseList(source,
                new OfficeIMO.Html.Css.HtmlCssSelectorOptions { Namespaces = namespaceContext }).IsSupported;
        } catch (OfficeIMO.Html.Css.HtmlCssSelectorLimitException) { return false; }
    }

    private static bool CanUseHybridSelectors(
        IReadOnlyList<string> selectors,
        OfficeIMO.Html.Css.HtmlCssNamespaceContext namespaceContext) {
        if (selectors.Count == 0) return false;
        try {
            foreach (string source in selectors) {
                string selector = TryParsePseudoElementSelector(source, out string host, out _) ? host : source;
                if (OfficeIMO.Html.Css.HtmlCssSelectorParser.ParseHybrid(selector,
                        new OfficeIMO.Html.Css.HtmlCssSelectorOptions { Namespaces = namespaceContext }).Selector == null) return false;
            }
            return true;
        } catch (OfficeIMO.Html.Css.HtmlCssSelectorLimitException) { return false; }
    }

    private static Dictionary<string, StyleDeclaration>? TryCreateOwnedDeclarations(
        OfficeIMO.Html.Css.HtmlCssQualifiedRule? rule) {
        if (rule?.Block?.IsClosed != true
            || rule.Contents.Any(node => node is OfficeIMO.Html.Css.HtmlCssInvalidSyntax)) return null;
        return TryCreateOwnedDeclarations(rule.Declarations);
    }

    private static Dictionary<string, StyleDeclaration>? TryCreateOwnedDeclarations(
        IReadOnlyList<OfficeIMO.Html.Css.HtmlCssDeclaration> sourceDeclarations) {
        var declarations = new Dictionary<string, StyleDeclaration>(HtmlCssPropertyNameComparer.Instance);
        int order = 0;
        foreach (OfficeIMO.Html.Css.HtmlCssDeclaration declaration in sourceDeclarations) {
            string propertyName = RestoreFontShorthandName(declaration.Name);
            if (!propertyName.StartsWith("--", StringComparison.Ordinal)
                && !OfficeIMO.Html.Css.HtmlCssPropertyCatalog.TryGet(propertyName, out _)) return null;
            OfficeIMO.Html.Css.HtmlCssPropertyParseResult parsed = OfficeIMO.Html.Css.HtmlCssPropertyParser.Parse(
                declaration, UnboundedPropertyTokenization);
            string value = RestoreProtectedDeclarationValue(StripCssCommentsOutsideStrings(parsed.AuthoredValue).Trim());
            if (parsed.Definition != null && !parsed.IsAccepted) return null;
            if (value.Length == 0 || !IsSupportedDeclarationValue(propertyName, value)) return null;
            if (string.Equals(propertyName, "color", StringComparison.OrdinalIgnoreCase)
                && (parsed.Value?.Kind == OfficeIMO.Html.Css.HtmlCssPropertyValueKind.NamedColor
                    || parsed.Value?.Kind == OfficeIMO.Html.Css.HtmlCssPropertyValueKind.HexColor)
                && OfficeIMO.Drawing.OfficeColor.TryParseCss(value, out OfficeIMO.Drawing.OfficeColor color)) {
                value = FormatComputedColor(color, color.A / 255D);
            }
            var candidate = new StyleDeclaration(propertyName, value, declaration.IsImportant) { DeclarationOrder = order++ };
            if (declarations.TryGetValue(propertyName, out StyleDeclaration? existing)
                && existing.IsImportant && !candidate.IsImportant) continue;
            declarations[propertyName] = candidate;
        }
        return declarations;
    }
}
