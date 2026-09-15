namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private sealed class StyleDeclaration {
        internal StyleDeclaration(string propertyName, string value, bool isImportant) {
            Value = value;
            IsImportant = isImportant;
            IsSupported = IsSupportedDeclarationValue(propertyName, value);
        }

        internal string Value { get; }
        internal bool IsImportant { get; }
        internal bool IsSupported { get; }
        internal int DeclarationOrder { get; set; }
    }

    private sealed class CascadedProperty {
        internal CascadedProperty(
            string value,
            bool isImportant,
            Specificity specificity,
            int order,
            CascadeLayerOrder? layerOrder = null,
            IEnumerable<CascadedProperty>? alternatives = null,
            bool inheritsComputedValue = false,
            int declarationOrder = 0,
            bool deferredFontShorthand = false,
            string? authoredValue = null,
            OfficeIMO.Html.Css.HtmlCssCascadeSourceKind source = OfficeIMO.Html.Css.HtmlCssCascadeSourceKind.StyleRule,
            string? selector = null,
            string? layerName = null) {
            Value = value;
            AuthoredValue = authoredValue ?? value;
            HasValue = true;
            IsImportant = isImportant;
            Specificity = specificity;
            Order = order;
            DeclarationOrder = declarationOrder;
            LayerOrder = layerOrder;
            Alternatives = MaterializeAlternatives(alternatives);
            InheritsComputedValue = inheritsComputedValue;
            IsDeferredFontShorthand = deferredFontShorthand;
            Source = source;
            Selector = selector;
            LayerName = layerName;
        }

        private CascadedProperty(bool isImportant, Specificity specificity, int order, CascadeLayerOrder? layerOrder,
            IEnumerable<CascadedProperty>? alternatives, bool revertsLayer, int declarationOrder, string authoredValue,
            OfficeIMO.Html.Css.HtmlCssCascadeSourceKind source, string? selector, string? layerName) {
            Value = string.Empty;
            AuthoredValue = authoredValue;
            HasValue = false;
            IsImportant = isImportant;
            Specificity = specificity;
            Order = order;
            DeclarationOrder = declarationOrder;
            LayerOrder = layerOrder;
            Alternatives = MaterializeAlternatives(alternatives);
            RevertsLayer = revertsLayer;
            InheritsComputedValue = false;
            Source = source;
            Selector = selector;
            LayerName = layerName;
        }

        internal static CascadedProperty Clear(bool isImportant, Specificity specificity, int order, CascadeLayerOrder? layerOrder,
            IEnumerable<CascadedProperty>? alternatives, int declarationOrder = 0, string authoredValue = "",
            OfficeIMO.Html.Css.HtmlCssCascadeSourceKind source = OfficeIMO.Html.Css.HtmlCssCascadeSourceKind.StyleRule,
            string? selector = null, string? layerName = null) {
            return new CascadedProperty(isImportant, specificity, order, layerOrder, alternatives, revertsLayer: false,
                declarationOrder, authoredValue, source, selector, layerName);
        }

        internal static CascadedProperty RevertLayer(bool isImportant, Specificity specificity, int order, CascadeLayerOrder? layerOrder,
            IEnumerable<CascadedProperty>? alternatives, int declarationOrder = 0, string authoredValue = "revert-layer",
            OfficeIMO.Html.Css.HtmlCssCascadeSourceKind source = OfficeIMO.Html.Css.HtmlCssCascadeSourceKind.StyleRule,
            string? selector = null, string? layerName = null) =>
            new CascadedProperty(isImportant, specificity, order, layerOrder, alternatives, revertsLayer: true,
                declarationOrder, authoredValue, source, selector, layerName);

        internal string Value { get; }
        internal string AuthoredValue { get; }
        internal bool HasValue { get; }
        internal bool IsImportant { get; }
        internal Specificity Specificity { get; }
        internal int Order { get; }
        internal int DeclarationOrder { get; }
        internal CascadeLayerOrder? LayerOrder { get; }
        internal IReadOnlyList<CascadedProperty> Alternatives { get; }
        internal bool RevertsLayer { get; }
        internal bool InheritsComputedValue { get; }
        internal bool IsDeferredFontShorthand { get; }
        internal OfficeIMO.Html.Css.HtmlCssCascadeSourceKind Source { get; }
        internal string? Selector { get; }
        internal string? LayerName { get; }

        internal CascadedProperty WithAlternative(CascadedProperty alternative) {
            var alternatives = new List<CascadedProperty>(Alternatives) { alternative };
            return RevertsLayer
                ? RevertLayer(IsImportant, Specificity, Order, LayerOrder, alternatives, DeclarationOrder, AuthoredValue, Source, Selector, LayerName)
                : HasValue
                    ? new CascadedProperty(Value, IsImportant, Specificity, Order, LayerOrder, alternatives, InheritsComputedValue,
                        DeclarationOrder, IsDeferredFontShorthand, AuthoredValue, Source, Selector, LayerName)
                    : Clear(IsImportant, Specificity, Order, LayerOrder, alternatives, DeclarationOrder, AuthoredValue, Source, Selector, LayerName);
        }

        private static IReadOnlyList<CascadedProperty> MaterializeAlternatives(IEnumerable<CascadedProperty>? alternatives) =>
            alternatives switch {
                null => Array.Empty<CascadedProperty>(),
                IReadOnlyList<CascadedProperty> list => list,
                _ => alternatives.ToArray()
            };
    }

    private readonly struct CssKeywordResolution {
        private CssKeywordResolution(bool hasValue, string value, bool inheritsComputedValue = false) {
            HasValue = hasValue;
            Value = value;
            InheritsComputedValue = inheritsComputedValue;
        }

        internal static CssKeywordResolution Clear => new CssKeywordResolution(false, string.Empty);
        internal static CssKeywordResolution ForValue(string value) => new CssKeywordResolution(true, value);
        internal static CssKeywordResolution ForInheritedValue(string value) => new CssKeywordResolution(true, value, inheritsComputedValue: true);

        internal bool HasValue { get; }
        internal string Value { get; }
        internal bool InheritsComputedValue { get; }
    }

    private sealed class Specificity {
        internal Specificity(int ids, int classesAttributesAndPseudoClasses, int elements) {
            Ids = ids;
            ClassesAttributesAndPseudoClasses = classesAttributesAndPseudoClasses;
            Elements = elements;
        }

        internal int Ids { get; }
        internal int ClassesAttributesAndPseudoClasses { get; }
        internal int Elements { get; }
        internal static Specificity Inherited { get; } = new Specificity(-1, -1, -1);
        internal static Specificity PresentationalHint { get; } = new Specificity(0, 0, 0);
        internal static Specificity Inline { get; } = new Specificity(int.MaxValue, int.MaxValue, int.MaxValue);

        internal int CompareTo(Specificity other) {
            if (Ids != other.Ids) {
                return Ids.CompareTo(other.Ids);
            }

            if (ClassesAttributesAndPseudoClasses != other.ClassesAttributesAndPseudoClasses) {
                return ClassesAttributesAndPseudoClasses.CompareTo(other.ClassesAttributesAndPseudoClasses);
            }

            return Elements.CompareTo(other.Elements);
        }
    }

    private sealed class StyleRule {
        internal StyleRule(
            string selector,
            Specificity specificity,
            int order,
            IDictionary<string, StyleDeclaration> declarations,
            CascadeLayerOrder? layerOrder = null,
            string? layerName = null,
            IEnumerable<ContainerRuleCondition>? containerConditions = null,
            OfficeIMO.Html.Css.HtmlCssNamespaceContext? namespaceContext = null,
            AngleSharp.Css.Dom.ISelector? providerSelector = null) {
            Selector = selector;
            string ownedSource = TryParsePseudoElementSelector(selector, out string hostSelector, out _) ? hostSelector : selector;
            OfficeIMO.Html.Css.HtmlCssSelector? ownedSelector = null;
            try {
                var selectorOptions = new OfficeIMO.Html.Css.HtmlCssSelectorOptions { Namespaces = namespaceContext };
                ownedSelector = OfficeIMO.Html.Css.HtmlCssSelectorParser.Parse(ownedSource, selectorOptions).Selector;
                if (ownedSelector == null) {
                    ownedSelector = OfficeIMO.Html.Css.HtmlCssSelectorParser.ParseHybrid(ownedSource, selectorOptions).Selector;
                }
            } catch (OfficeIMO.Html.Css.HtmlCssSelectorLimitException) {
                // The conversion stylesheet budget remains authoritative. Selectors outside the
                // standalone parser budget stay on the retained provider path.
            }
            OwnedSelector = ownedSelector;
            ProviderSelector = ownedSelector == null || ownedSelector.RequiresProviderMatching ? providerSelector : null;
            Specificity = ownedSelector != null && !ownedSelector.RequiresProviderMatching
                ? new Specificity(ownedSelector.Specificity.Ids, ownedSelector.Specificity.Classes, ownedSelector.Specificity.Types)
                : ProviderSelector != null
                    ? new Specificity(ProviderSelector.Specificity.Ids, ProviderSelector.Specificity.Classes, ProviderSelector.Specificity.Tags)
                    : specificity;
            Order = order;
            Declarations = new Dictionary<string, StyleDeclaration>(declarations, HtmlCssPropertyNameComparer.Instance);
            LayerOrder = layerOrder;
            LayerName = layerName;
            ContainerConditions = new List<ContainerRuleCondition>(containerConditions ?? Array.Empty<ContainerRuleCondition>()).AsReadOnly();
            CandidateKey = GetSelectorCandidateKey(selector);
        }

        internal string Selector { get; }
        internal OfficeIMO.Html.Css.HtmlCssSelector? OwnedSelector { get; }
        internal AngleSharp.Css.Dom.ISelector? ProviderSelector { get; }
        internal Specificity Specificity { get; }
        internal int Order { get; }
        internal IReadOnlyDictionary<string, StyleDeclaration> Declarations { get; }
        internal CascadeLayerOrder? LayerOrder { get; }
        internal string? LayerName { get; }
        internal IReadOnlyList<ContainerRuleCondition> ContainerConditions { get; }
        internal SelectorCandidateKey CandidateKey { get; }
    }

    private enum SelectorCandidateKind {
        Universal,
        Tag,
        Class,
        Id
    }

    private readonly struct SelectorCandidateKey {
        internal SelectorCandidateKey(SelectorCandidateKind kind, string value) {
            Kind = kind;
            Value = value;
        }

        internal SelectorCandidateKind Kind { get; }
        internal string Value { get; }
    }

    /// <summary>
    /// Indexes each selector by one required token from its rightmost compound. Rules that cannot
    /// be classified conservatively stay universal, so indexing changes work performed rather
    /// than CSS semantics.
    /// </summary>
    private sealed class StyleRuleIndex {
        private readonly List<StyleRule> _universal = new List<StyleRule>();
        private readonly Dictionary<string, List<StyleRule>> _tags = new Dictionary<string, List<StyleRule>>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, List<StyleRule>> _classes = new Dictionary<string, List<StyleRule>>(StringComparer.Ordinal);
        private readonly Dictionary<string, List<StyleRule>> _ids = new Dictionary<string, List<StyleRule>>(StringComparer.Ordinal);

        internal StyleRuleIndex(
            IEnumerable<StyleRule> rules,
            IReadOnlyDictionary<string, CustomPropertyRegistration>? customPropertyRegistrations = null) {
            CustomPropertyRegistrations = customPropertyRegistrations
                ?? new Dictionary<string, CustomPropertyRegistration>(HtmlCssPropertyNameComparer.Instance);
            foreach (StyleRule rule in rules) {
                switch (rule.CandidateKey.Kind) {
                    case SelectorCandidateKind.Tag:
                        Add(_tags, rule.CandidateKey.Value, rule);
                        break;
                    case SelectorCandidateKind.Class:
                        Add(_classes, rule.CandidateKey.Value, rule);
                        break;
                    case SelectorCandidateKind.Id:
                        Add(_ids, rule.CandidateKey.Value, rule);
                        break;
                    default:
                        _universal.Add(rule);
                        break;
                }
            }
        }

        internal IReadOnlyDictionary<string, CustomPropertyRegistration> CustomPropertyRegistrations { get; }

        internal IReadOnlyList<StyleRule> GetCandidates(AngleSharp.Dom.IElement element) {
            var candidates = new List<StyleRule>(_universal.Count + 8);
            candidates.AddRange(_universal);
            AddMatches(_tags, element.LocalName ?? element.TagName ?? string.Empty, candidates);
            string? id = element.Id;
            if (!string.IsNullOrEmpty(id)) AddMatches(_ids, id!, candidates);
            foreach (string className in element.ClassList) AddMatches(_classes, className, candidates);
            if (candidates.Count > 1) candidates.Sort((left, right) => left.Order.CompareTo(right.Order));
            return candidates;
        }

        private static void Add(Dictionary<string, List<StyleRule>> index, string key, StyleRule rule) {
            if (!index.TryGetValue(key, out List<StyleRule>? rules)) {
                rules = new List<StyleRule>();
                index[key] = rules;
            }
            rules.Add(rule);
        }

        private static void AddMatches(
            Dictionary<string, List<StyleRule>> index,
            string key,
            ICollection<StyleRule> candidates) {
            if (index.TryGetValue(key, out List<StyleRule>? rules)) {
                foreach (StyleRule rule in rules) candidates.Add(rule);
            }
        }
    }

    private sealed class CustomPropertyRegistration {
        internal CustomPropertyRegistration(string name, string syntax, bool inherits, string? initialValue) {
            Name = name;
            Syntax = syntax;
            Inherits = inherits;
            InitialValue = initialValue;
        }

        internal string Name { get; }
        internal string Syntax { get; }
        internal bool Inherits { get; }
        internal string? InitialValue { get; }
    }

    private sealed class ContainerRuleCondition {
        internal ContainerRuleCondition(string name, string condition) {
            Name = name;
            Condition = condition;
        }

        internal string Name { get; }
        internal string Condition { get; }
    }

    private sealed class ContainerQueryContext {
        internal ContainerQueryContext(
            IReadOnlyList<string> names,
            string type,
            double width,
            double? height,
            double fontSize,
            double inheritedFontSize,
            double rootFontSize,
            IReadOnlyDictionary<string, string> properties) {
            Names = names;
            Type = type;
            Width = width;
            Height = height;
            FontSize = fontSize;
            InheritedFontSize = inheritedFontSize;
            RootFontSize = rootFontSize;
            Properties = properties;
        }

        internal IReadOnlyList<string> Names { get; }
        internal string Type { get; }
        internal double Width { get; }
        internal double? Height { get; }
        internal double FontSize { get; }
        internal double InheritedFontSize { get; }
        internal double RootFontSize { get; }
        internal IReadOnlyDictionary<string, string> Properties { get; }
    }

}
