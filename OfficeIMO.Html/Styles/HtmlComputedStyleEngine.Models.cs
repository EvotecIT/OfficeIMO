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
        internal CascadedProperty(string value, bool isImportant, Specificity specificity, int order) {
            Value = value;
            HasValue = true;
            IsImportant = isImportant;
            Specificity = specificity;
            Order = order;
        }

        private CascadedProperty(bool isImportant, Specificity specificity, int order) {
            Value = string.Empty;
            HasValue = false;
            IsImportant = isImportant;
            Specificity = specificity;
            Order = order;
        }

        internal static CascadedProperty Clear(bool isImportant, Specificity specificity, int order) {
            return new CascadedProperty(isImportant, specificity, order);
        }

        internal string Value { get; }
        internal bool HasValue { get; }
        internal bool IsImportant { get; }
        internal Specificity Specificity { get; }
        internal int Order { get; }
    }

    private readonly struct CssKeywordResolution {
        private CssKeywordResolution(bool hasValue, string value) {
            HasValue = hasValue;
            Value = value;
        }

        internal static CssKeywordResolution Clear => new CssKeywordResolution(false, string.Empty);
        internal static CssKeywordResolution ForValue(string value) => new CssKeywordResolution(true, value);

        internal bool HasValue { get; }
        internal string Value { get; }
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
        internal StyleRule(string selector, Specificity specificity, int order, IDictionary<string, StyleDeclaration> declarations) {
            Selector = selector;
            Specificity = specificity;
            Order = order;
            Declarations = new Dictionary<string, StyleDeclaration>(declarations, StringComparer.OrdinalIgnoreCase);
            CandidateKey = GetSelectorCandidateKey(selector);
        }

        internal string Selector { get; }
        internal Specificity Specificity { get; }
        internal int Order { get; }
        internal IReadOnlyDictionary<string, StyleDeclaration> Declarations { get; }
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
}
