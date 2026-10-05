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
        internal CascadedProperty(string value, bool isImportant, Specificity specificity, int order, CascadeLayerOrder? layerOrder = null, IEnumerable<CascadedProperty>? alternatives = null, bool inheritsComputedValue = false, int declarationOrder = 0, bool deferredFontShorthand = false, string? deferredGridShorthand = null) {
            Value = value;
            HasValue = true;
            IsImportant = isImportant;
            Specificity = specificity;
            Order = order;
            DeclarationOrder = declarationOrder;
            LayerOrder = layerOrder;
            Alternatives = MaterializeAlternatives(alternatives);
            InheritsComputedValue = inheritsComputedValue;
            IsDeferredFontShorthand = deferredFontShorthand;
            DeferredGridShorthand = deferredGridShorthand;
        }

        private CascadedProperty(bool isImportant, Specificity specificity, int order, CascadeLayerOrder? layerOrder, IEnumerable<CascadedProperty>? alternatives, bool revertsLayer, int declarationOrder) {
            Value = string.Empty;
            HasValue = false;
            IsImportant = isImportant;
            Specificity = specificity;
            Order = order;
            DeclarationOrder = declarationOrder;
            LayerOrder = layerOrder;
            Alternatives = MaterializeAlternatives(alternatives);
            RevertsLayer = revertsLayer;
            InheritsComputedValue = false;
        }

        internal static CascadedProperty Clear(bool isImportant, Specificity specificity, int order, CascadeLayerOrder? layerOrder, IEnumerable<CascadedProperty>? alternatives, int declarationOrder = 0) {
            return new CascadedProperty(isImportant, specificity, order, layerOrder, alternatives, revertsLayer: false, declarationOrder);
        }

        internal static CascadedProperty RevertLayer(bool isImportant, Specificity specificity, int order, CascadeLayerOrder? layerOrder, IEnumerable<CascadedProperty>? alternatives, int declarationOrder = 0) =>
            new CascadedProperty(isImportant, specificity, order, layerOrder, alternatives, revertsLayer: true, declarationOrder);

        internal string Value { get; }
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
        internal string? DeferredGridShorthand { get; }

        internal CascadedProperty WithAlternative(CascadedProperty alternative) {
            var alternatives = new List<CascadedProperty>(Alternatives) { alternative };
            return RevertsLayer
                ? RevertLayer(IsImportant, Specificity, Order, LayerOrder, alternatives, DeclarationOrder)
                : HasValue
                    ? new CascadedProperty(Value, IsImportant, Specificity, Order, LayerOrder, alternatives, InheritsComputedValue, DeclarationOrder, IsDeferredFontShorthand, DeferredGridShorthand)
                    : Clear(IsImportant, Specificity, Order, LayerOrder, alternatives, DeclarationOrder);
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
            IEnumerable<ContainerRuleCondition>? containerConditions = null) {
            Selector = selector;
            Specificity = specificity;
            Order = order;
            Declarations = new Dictionary<string, StyleDeclaration>(declarations, HtmlCssPropertyNameComparer.Instance);
            LayerOrder = layerOrder;
            ContainerConditions = new List<ContainerRuleCondition>(containerConditions ?? Array.Empty<ContainerRuleCondition>()).AsReadOnly();
            CandidateKey = GetSelectorCandidateKey(selector);
            if (TryParsePseudoElementSelector(selector, out string hostSelector, out HtmlPseudoElementKind kind)) {
                MatchingSelector = hostSelector;
                PseudoKind = kind;
            } else {
                MatchingSelector = selector;
            }
        }

        internal string Selector { get; }
        internal string MatchingSelector { get; }
        internal HtmlPseudoElementKind? PseudoKind { get; }
        internal Specificity Specificity { get; }
        internal int Order { get; }
        internal IReadOnlyDictionary<string, StyleDeclaration> Declarations { get; }
        internal CascadeLayerOrder? LayerOrder { get; }
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
