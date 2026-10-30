using AngleSharp.Dom;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    /// <summary>Reuses declaration resolution after matching and its operation limits have run for this element.</summary>
    private static HtmlComputedStyle ResolveElementComputedStyle(
        IElement element,
        Dictionary<string, CascadedProperty> properties,
        IReadOnlyDictionary<string, string>? parentProperties,
        StyleRuleIndex rules,
        MediaEnvironment environment,
        bool includeCascadeTraces,
        bool enforceResolutionLimits) {
        ComputedStyleReuseKey? key = rules.ComputedReuse.CreateKey(element, properties, parentProperties, includeCascadeTraces);
        if (key != null && rules.ComputedReuse.TryGet(key, out HtmlComputedStyle? template))
            return template.CreateElementSnapshot();

        Dictionary<string, string> resolvedProperties = ResolveComputedProperties(properties, parentProperties,
            out HashSet<string> inheritedProperties, out HashSet<string> resetProperties,
            out HashSet<string> originRevertedProperties,
            out HashSet<string> specifiedProperties,
            out Dictionary<string, HtmlCssCascadePriority> cascadePriorities,
            out Dictionary<string, OfficeIMO.Html.Css.HtmlCssCascadeTrace>? cascadeTraces,
            includeCascadeTraces,
            rules.CustomPropertyRegistrations,
            enforceResolutionLimits);
        if (!resolvedProperties.ContainsKey("display")) resolvedProperties["display"] = HtmlElementDisplay.GetDefaultValue(element);
        MaterializeInitialFontValues(resolvedProperties, resetProperties, originRevertedProperties, cascadeTraces, environment.FontDefaults);
        ApplyHtmlWhiteSpaceDefaults(element, properties.ContainsKey("white-space"), resolvedProperties,
            inheritedProperties, resetProperties, originRevertedProperties, cascadePriorities, cascadeTraces);
        ApplyHtmlFontDefaults(element, properties, parentProperties, resolvedProperties,
            inheritedProperties, resetProperties, originRevertedProperties, cascadePriorities, cascadeTraces, environment.FontDefaults);
        ApplyHtmlTableAlignmentDefault(element, properties.ContainsKey("text-align"), parentProperties, resolvedProperties,
            inheritedProperties, resetProperties, originRevertedProperties, cascadePriorities, cascadeTraces);
        HtmlComputedStyle style = HtmlComputedStyle.FromOwnedCollections(
            resolvedProperties, inheritedProperties, resetProperties, originRevertedProperties,
            specifiedProperties, cascadePriorities, cascadeTraces);
        if (key == null) return style;
        rules.ComputedReuse.Add(key, style);
        return style.CreateElementSnapshot();
    }

    /// <summary>Owns successful immutable declaration payloads for one computation; traced and generated contexts use their ordinary path.</summary>
    private sealed class ComputedStyleReuseCache {
        private const int MaximumEntries = 128;
        private const int MaximumDeclarations = 64;
        private const int MaximumCandidateNodes = 256;
        private const int MaximumCandidateDepth = 16;
        private const int MaximumSourceCharacters = 8192;
        private readonly Dictionary<ComputedStyleReuseKey, HtmlComputedStyle> _styles = new();

        internal ComputedStyleReuseKey? CreateKey(IElement element,
            IReadOnlyDictionary<string, CascadedProperty> properties,
            IReadOnlyDictionary<string, string>? parentProperties, bool includeCascadeTraces) {
            if (includeCascadeTraces || _styles.Count >= MaximumEntries || properties.Count > MaximumDeclarations) return null;
            long characters = 0;
            int nodes = 0;
            int hash = 17;
            foreach (KeyValuePair<string, CascadedProperty> pair in properties) {
                characters += pair.Key.Length;
                if (!TryHashDeclaration(pair.Value, 0, ref nodes, ref characters, out int declarationHash)) return null;
                hash = unchecked((hash * 31 + StringComparer.Ordinal.GetHashCode(pair.Key)) * 31 + declarationHash);
            }
            // Preserve enumeration order as well as winning and losing declaration
            // provenance. The captured candidates are private immutable owner objects.
            return new ComputedStyleReuseKey(parentProperties, element.NamespaceUri ?? string.Empty,
                element.LocalName, HtmlElementDisplay.GetDefaultValue(element),
                element.LocalName == "wbr" && HasHtmlNoBreakAncestor(element), properties.ToArray(), hash);
        }

        private static bool TryHashDeclaration(CascadedProperty value, int depth, ref int nodes,
            ref long characters, out int hash) {
            hash = 0;
            if (depth > MaximumCandidateDepth || ++nodes > MaximumCandidateNodes) return false;
            characters += value.Value.Length;
            characters += value.AuthoredValue.Length;
            characters += value.DeferredLayoutShorthand?.Length ?? 0;
            characters += value.Selector?.Length ?? 0;
            characters += value.LayerName?.Length ?? 0;
            if (characters > MaximumSourceCharacters) return false;
            unchecked {
                hash = ToCascadePriority(value).GetHashCode();
                hash = hash * 31 + StringComparer.Ordinal.GetHashCode(value.Value);
                hash = hash * 31 + StringComparer.Ordinal.GetHashCode(value.AuthoredValue);
                hash = hash * 31 + value.HasValue.GetHashCode();
                hash = hash * 31 + value.RevertsLayer.GetHashCode();
                hash = hash * 31 + value.InheritsComputedValue.GetHashCode();
                hash = hash * 31 + value.IsDeferredFontShorthand.GetHashCode();
                hash = hash * 31 + (value.DeferredLayoutShorthand?.GetHashCode() ?? 0);
                hash = hash * 31 + (int)value.Source;
                hash = hash * 31 + (value.Selector?.GetHashCode() ?? 0);
                hash = hash * 31 + (value.LayerName?.GetHashCode() ?? 0);
                hash = hash * 31 + ReferenceEquals(value.Specificity, Specificity.Inherited).GetHashCode();
                hash = hash * 31 + ReferenceEquals(value.Specificity, Specificity.Inline).GetHashCode();
                hash = hash * 31 + ReferenceEquals(value.Specificity, Specificity.PresentationalHint).GetHashCode();
                hash = hash * 31 + value.Alternatives.Count;
                foreach (CascadedProperty alternative in value.Alternatives) {
                    if (!TryHashDeclaration(alternative, depth + 1, ref nodes, ref characters, out int alternativeHash)) return false;
                    hash = hash * 31 + alternativeHash;
                }
            }
            return true;
        }

        internal bool TryGet(ComputedStyleReuseKey key, out HtmlComputedStyle template) => _styles.TryGetValue(key, out template!);
        internal void Add(ComputedStyleReuseKey key, HtmlComputedStyle template) => _styles.Add(key, template);
    }

    private sealed class ComputedStyleReuseKey : IEquatable<ComputedStyleReuseKey> {
        private readonly IReadOnlyDictionary<string, string>? _parent;
        private readonly string _namespace;
        private readonly string _tag;
        private readonly string _display;
        private readonly bool _noBreakAncestor;
        private readonly KeyValuePair<string, CascadedProperty>[] _declarations;
        private readonly int _hashCode;

        internal ComputedStyleReuseKey(IReadOnlyDictionary<string, string>? parent, string elementNamespace,
            string tag, string display, bool noBreakAncestor,
            KeyValuePair<string, CascadedProperty>[] declarations, int declarationHash) {
            _parent = parent;
            _namespace = elementNamespace;
            _tag = tag;
            _display = display;
            _noBreakAncestor = noBreakAncestor;
            _declarations = declarations;
            unchecked {
                int hash = parent == null ? 0 : RuntimeHelpers.GetHashCode(parent);
                hash = hash * 31 + StringComparer.Ordinal.GetHashCode(elementNamespace);
                hash = hash * 31 + StringComparer.Ordinal.GetHashCode(tag);
                hash = hash * 31 + StringComparer.Ordinal.GetHashCode(display);
                hash = hash * 31 + noBreakAncestor.GetHashCode();
                _hashCode = hash * 31 + declarationHash;
            }
        }

        public bool Equals(ComputedStyleReuseKey? other) {
            if (ReferenceEquals(this, other)) return true;
            if (other == null || _hashCode != other._hashCode || !ReferenceEquals(_parent, other._parent)
                || _namespace != other._namespace || _tag != other._tag || _display != other._display
                || _noBreakAncestor != other._noBreakAncestor || _declarations.Length != other._declarations.Length) return false;
            for (int index = 0; index < _declarations.Length; index++) {
                if (_declarations[index].Key != other._declarations[index].Key
                    || !SameDeclaration(_declarations[index].Value, other._declarations[index].Value)) return false;
            }
            return true;
        }

        // Creation already bounded both candidate trees. Preserve every field
        // and the exact ordered rollback graph without reflective struct equality.
        private static bool SameDeclaration(CascadedProperty left, CascadedProperty right) {
            if (ReferenceEquals(left, right)) return true;
            if (!ToCascadePriority(left).Equals(ToCascadePriority(right)) || left.Value != right.Value
                || left.AuthoredValue != right.AuthoredValue || left.HasValue != right.HasValue
                || left.RevertsLayer != right.RevertsLayer || left.InheritsComputedValue != right.InheritsComputedValue
                || left.IsDeferredFontShorthand != right.IsDeferredFontShorthand
                || left.DeferredLayoutShorthand != right.DeferredLayoutShorthand || left.Source != right.Source
                || left.Selector != right.Selector || left.LayerName != right.LayerName
                || ReferenceEquals(left.Specificity, Specificity.Inherited) != ReferenceEquals(right.Specificity, Specificity.Inherited)
                || ReferenceEquals(left.Specificity, Specificity.Inline) != ReferenceEquals(right.Specificity, Specificity.Inline)
                || ReferenceEquals(left.Specificity, Specificity.PresentationalHint) != ReferenceEquals(right.Specificity, Specificity.PresentationalHint)
                || left.Alternatives.Count != right.Alternatives.Count) return false;
            using IEnumerator<CascadedProperty> leftCandidates = left.Alternatives.GetEnumerator();
            using IEnumerator<CascadedProperty> rightCandidates = right.Alternatives.GetEnumerator();
            while (leftCandidates.MoveNext()) {
                rightCandidates.MoveNext();
                if (!SameDeclaration(leftCandidates.Current, rightCandidates.Current)) return false;
            }
            return true;
        }

        public override bool Equals(object? obj) => obj is ComputedStyleReuseKey other && Equals(other);
        public override int GetHashCode() => _hashCode;
    }
}
