using OfficeIMO.Html.Css;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private static Dictionary<string, HtmlCssCascadeTrace> BuildCascadeTraces(
        IReadOnlyDictionary<string, CascadedProperty> properties,
        IReadOnlyDictionary<string, string> resolved,
        ISet<string> inherited,
        ISet<string> reset,
        ISet<string> invalidAtComputedValue) {
        var traces = new Dictionary<string, HtmlCssCascadeTrace>(HtmlCssPropertyNameComparer.Instance);
        foreach (HtmlCssPropertyDefinition definition in HtmlCssPropertyCatalog.All) {
            string name = definition.Name;
            resolved.TryGetValue(name, out string? computedValue);
            bool isInherited = inherited.Contains(name);
            bool isReset = reset.Contains(name) || invalidAtComputedValue.Contains(name) && !isInherited;
            if (!properties.TryGetValue(name, out CascadedProperty? root)) {
                if (isInherited && computedValue != null) {
                    traces[name] = new HtmlCssCascadeTrace(name, computedValue, true, false, new[] {
                        new HtmlCssCascadeCandidate(computedValue, computedValue, HtmlCssCascadeSourceKind.Inherited,
                            HtmlCssCascadeDecision.Inherited, false,
                            HtmlCssPropertyParser.Parse(name, computedValue, UnboundedPropertyTokenization).Status,
                            false, new HtmlCssSpecificity(-1, -1, -1), -1, -1)
                    });
                }
                continue;
            }

            CascadedProperty? effective = ResolveLayerRevert(root);
            var retained = new List<CascadedProperty> { root };
            retained.AddRange(root.Alternatives);
            retained.Sort((left, right) => {
                int rule = left.Order.CompareTo(right.Order);
                return rule != 0 ? rule : left.DeclarationOrder.CompareTo(right.DeclarationOrder);
            });
            var candidates = new List<HtmlCssCascadeCandidate>(retained.Count);
            foreach (CascadedProperty candidate in retained) {
                HtmlCssCascadeDecision decision;
                bool isWinningDeclaration = ReferenceEquals(candidate, root);
                if (!isWinningDeclaration && !ReferenceEquals(candidate, effective))
                    decision = HtmlCssCascadeDecision.Overridden;
                else if (isWinningDeclaration && candidate.RevertsLayer) decision = HtmlCssCascadeDecision.RevertedLayer;
                else if (isWinningDeclaration && string.Equals(candidate.AuthoredValue.Trim(), "revert", StringComparison.OrdinalIgnoreCase))
                    decision = HtmlCssCascadeDecision.RevertedOrigin;
                else if (isWinningDeclaration && invalidAtComputedValue.Contains(name))
                    decision = HtmlCssCascadeDecision.InvalidAtComputedValue;
                else if (isWinningDeclaration && (!candidate.HasValue || IsResetKeyword(candidate.AuthoredValue, definition.IsInherited)))
                    decision = HtmlCssCascadeDecision.Reset;
                else if (ReferenceEquals(candidate, effective)) decision = HtmlCssCascadeDecision.Selected;
                else decision = HtmlCssCascadeDecision.Overridden;
                if (ReferenceEquals(candidate, effective) && decision == HtmlCssCascadeDecision.Reset) isReset = true;
                candidates.Add(new HtmlCssCascadeCandidate(
                    candidate.AuthoredValue,
                    candidate.HasValue ? candidate.Value : null,
                    candidate.Source,
                    decision,
                    isWinningDeclaration,
                    HtmlCssPropertyParser.Parse(name, candidate.AuthoredValue, UnboundedPropertyTokenization).Status,
                    candidate.IsImportant,
                    new HtmlCssSpecificity(candidate.Specificity.Ids,
                        candidate.Specificity.ClassesAttributesAndPseudoClasses, candidate.Specificity.Elements),
                    candidate.Order,
                    candidate.DeclarationOrder,
                    candidate.Selector,
                    candidate.LayerName));
            }
            traces[name] = new HtmlCssCascadeTrace(name, computedValue, isInherited, isReset, candidates);
        }
        return traces;
    }

    private static bool IsResetKeyword(string value, bool inheritedProperty) {
        string normalized = value.Trim();
        return string.Equals(normalized, "initial", StringComparison.OrdinalIgnoreCase)
            || !inheritedProperty && string.Equals(normalized, "unset", StringComparison.OrdinalIgnoreCase);
    }
}
