namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private sealed class StyleDeclaration {
        internal StyleDeclaration(string value, bool isImportant) {
            Value = value;
            IsImportant = isImportant;
        }

        internal string Value { get; }
        internal bool IsImportant { get; }
    }

    private sealed class CascadedProperty {
        internal CascadedProperty(string value, bool isImportant, Specificity specificity, int order, CascadeLayerOrder? layerOrder = null, IEnumerable<CascadedProperty>? alternatives = null, bool inheritsComputedValue = false) {
            Value = value;
            HasValue = true;
            IsImportant = isImportant;
            Specificity = specificity;
            Order = order;
            LayerOrder = layerOrder;
            Alternatives = new List<CascadedProperty>(alternatives ?? Array.Empty<CascadedProperty>()).AsReadOnly();
            InheritsComputedValue = inheritsComputedValue;
        }

        private CascadedProperty(bool isImportant, Specificity specificity, int order, CascadeLayerOrder? layerOrder, IEnumerable<CascadedProperty>? alternatives, bool revertsLayer) {
            Value = string.Empty;
            HasValue = false;
            IsImportant = isImportant;
            Specificity = specificity;
            Order = order;
            LayerOrder = layerOrder;
            Alternatives = new List<CascadedProperty>(alternatives ?? Array.Empty<CascadedProperty>()).AsReadOnly();
            RevertsLayer = revertsLayer;
            InheritsComputedValue = false;
        }

        internal static CascadedProperty Clear(bool isImportant, Specificity specificity, int order, CascadeLayerOrder? layerOrder, IEnumerable<CascadedProperty>? alternatives) {
            return new CascadedProperty(isImportant, specificity, order, layerOrder, alternatives, revertsLayer: false);
        }

        internal static CascadedProperty RevertLayer(bool isImportant, Specificity specificity, int order, CascadeLayerOrder? layerOrder, IEnumerable<CascadedProperty>? alternatives) =>
            new CascadedProperty(isImportant, specificity, order, layerOrder, alternatives, revertsLayer: true);

        internal string Value { get; }
        internal bool HasValue { get; }
        internal bool IsImportant { get; }
        internal Specificity Specificity { get; }
        internal int Order { get; }
        internal CascadeLayerOrder? LayerOrder { get; }
        internal IReadOnlyList<CascadedProperty> Alternatives { get; }
        internal bool RevertsLayer { get; }
        internal bool InheritsComputedValue { get; }

        internal CascadedProperty WithAlternative(CascadedProperty alternative) {
            var alternatives = new List<CascadedProperty>(Alternatives) { alternative };
            return RevertsLayer
                ? RevertLayer(IsImportant, Specificity, Order, LayerOrder, alternatives)
                : HasValue
                    ? new CascadedProperty(Value, IsImportant, Specificity, Order, LayerOrder, alternatives, InheritsComputedValue)
                    : Clear(IsImportant, Specificity, Order, LayerOrder, alternatives);
        }
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
            Declarations = new Dictionary<string, StyleDeclaration>(declarations, StringComparer.OrdinalIgnoreCase);
            LayerOrder = layerOrder;
            ContainerConditions = new List<ContainerRuleCondition>(containerConditions ?? Array.Empty<ContainerRuleCondition>()).AsReadOnly();
            CandidateKey = GetSelectorCandidateKey(selector);
        }

        internal string Selector { get; }
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

        internal StyleRuleIndex(IEnumerable<StyleRule> rules) {
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
            double rootFontSize,
            IReadOnlyDictionary<string, string> properties) {
            Names = names;
            Type = type;
            Width = width;
            Height = height;
            FontSize = fontSize;
            RootFontSize = rootFontSize;
            Properties = properties;
        }

        internal IReadOnlyList<string> Names { get; }
        internal string Type { get; }
        internal double Width { get; }
        internal double? Height { get; }
        internal double FontSize { get; }
        internal double RootFontSize { get; }
        internal IReadOnlyDictionary<string, string> Properties { get; }
    }

    private sealed class CascadeLayerRegistry {
        private readonly Dictionary<string, Dictionary<string, int>> _children = new Dictionary<string, Dictionary<string, int>>(StringComparer.Ordinal);
        private int _anonymous;

        internal CascadeLayerOrder Register(string name) {
            string[] components = name.Split(new[] { '.' }, StringSplitOptions.RemoveEmptyEntries);
            var positions = new List<int>(components.Length);
            string parent = string.Empty;
            foreach (string component in components) {
                if (!_children.TryGetValue(parent, out Dictionary<string, int>? children)) {
                    children = new Dictionary<string, int>(StringComparer.Ordinal);
                    _children[parent] = children;
                }
                if (!children.TryGetValue(component, out int position)) {
                    position = children.Count;
                    children[component] = position;
                }
                positions.Add(position);
                parent = parent.Length == 0 ? component : parent + "." + component;
            }
            return new CascadeLayerOrder(positions);
        }

        internal string RegisterAnonymous(string? parent) {
            string name = (parent == null ? string.Empty : parent + ".") + "#anonymous-" + (++_anonymous).ToString(System.Globalization.CultureInfo.InvariantCulture);
            Register(name);
            return name;
        }

        internal CascadeLayerOrder GetOrder(string name) => Register(name);
    }

    private sealed class CascadeLayerOrder : IEquatable<CascadeLayerOrder> {
        private readonly int[] _components;

        internal CascadeLayerOrder(IEnumerable<int> components) {
            _components = components.ToArray();
        }

        internal int CompareTo(CascadeLayerOrder other) {
            int shared = Math.Min(_components.Length, other._components.Length);
            for (int index = 0; index < shared; index++) {
                if (_components[index] != other._components[index]) {
                    return _components[index].CompareTo(other._components[index]);
                }
            }

            // Declarations directly in a layer have normal precedence over declarations
            // in its nested sublayers. Important declarations reverse this ordering later.
            return other._components.Length.CompareTo(_components.Length);
        }

        public bool Equals(CascadeLayerOrder? other) =>
            other != null && _components.SequenceEqual(other._components);

        public override bool Equals(object? obj) => Equals(obj as CascadeLayerOrder);

        public override int GetHashCode() {
            unchecked {
                int hash = 17;
                foreach (int component in _components) hash = (hash * 31) + component;
                return hash;
            }
        }
    }
}
