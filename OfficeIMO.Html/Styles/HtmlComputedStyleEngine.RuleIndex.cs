using AngleSharp.Dom;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    /// <summary>
    /// Separates element and pseudo-element rules before matching so each selector consumes
    /// budget only for its own target. Every bucket keeps the same conservative token index
    /// and source order, including rules whose host selector must stay universal.
    /// </summary>
    private sealed class StyleRuleIndex {
        private readonly SelectorRuleBucket _elements = new SelectorRuleBucket();
        private readonly Dictionary<HtmlPseudoElementKind, SelectorRuleBucket> _pseudoElements =
            new Dictionary<HtmlPseudoElementKind, SelectorRuleBucket>();

        internal StyleRuleIndex(
            IEnumerable<StyleRule> rules,
            IReadOnlyDictionary<string, CustomPropertyRegistration>? customPropertyRegistrations = null) {
            CustomPropertyRegistrations = customPropertyRegistrations
                ?? new Dictionary<string, CustomPropertyRegistration>(HtmlCssPropertyNameComparer.Instance);
            foreach (StyleRule rule in rules) {
                if (!rule.PseudoElementKind.HasValue) {
                    _elements.Add(rule);
                    continue;
                }

                HtmlPseudoElementKind kind = rule.PseudoElementKind.Value;
                if (!_pseudoElements.TryGetValue(kind, out SelectorRuleBucket? bucket)) {
                    bucket = new SelectorRuleBucket();
                    _pseudoElements[kind] = bucket;
                }
                bucket.Add(rule);
            }
        }

        internal IReadOnlyDictionary<string, CustomPropertyRegistration> CustomPropertyRegistrations { get; }

        internal IReadOnlyList<StyleRule> GetCandidates(IElement element) => _elements.GetCandidates(element);

        internal IReadOnlyList<StyleRule> GetPseudoCandidates(IElement element, HtmlPseudoElementKind kind) =>
            _pseudoElements.TryGetValue(kind, out SelectorRuleBucket? bucket)
                ? bucket.GetCandidates(element)
                : Array.Empty<StyleRule>();
    }

    /// <summary>Indexes each rule by one required token from its rightmost host compound.</summary>
    private sealed class SelectorRuleBucket {
        private readonly List<StyleRule> _universal = new List<StyleRule>();
        private readonly Dictionary<string, List<StyleRule>> _tags = new Dictionary<string, List<StyleRule>>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, List<StyleRule>> _classes = new Dictionary<string, List<StyleRule>>(StringComparer.Ordinal);
        private readonly Dictionary<string, List<StyleRule>> _ids = new Dictionary<string, List<StyleRule>>(StringComparer.Ordinal);

        internal void Add(StyleRule rule) {
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

        internal IReadOnlyList<StyleRule> GetCandidates(IElement element) {
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

        private static void AddMatches(Dictionary<string, List<StyleRule>> index, string key, ICollection<StyleRule> candidates) {
            if (index.TryGetValue(key, out List<StyleRule>? rules)) {
                foreach (StyleRule rule in rules) candidates.Add(rule);
            }
        }
    }
}
