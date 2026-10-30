using System.Runtime.CompilerServices;

namespace OfficeIMO.Html;

internal sealed partial class HtmlCssPageRuleSet {
    // The page-rule set belongs to one rendering operation. Font discovery,
    // metric preparation and repeated pages share computed cascades, while used
    // typography, geometry, counters and running content remain per-call work.
    private readonly Dictionary<MarginStyleContextKey, IReadOnlyList<MarginStyleContext>> _marginStyleContexts = new();
    internal int CachedMarginContextCount => _marginStyleContexts.Count;

    private readonly struct MarginStyleContext {
        internal MarginStyleContext(HtmlComputedStyle page, HtmlComputedStyle margin,
            HtmlCssPageMarginPosition position, HtmlCssGeneratedContentTemplate content) {
            Page = page;
            Margin = margin;
            Position = position;
            Content = content;
        }

        internal HtmlComputedStyle Page { get; }
        internal HtmlComputedStyle Margin { get; }
        internal HtmlCssPageMarginPosition Position { get; }
        internal HtmlCssGeneratedContentTemplate Content { get; }
    }

    /// <summary>Identifies an immutable cascade by its actual matched rules and parent snapshots.</summary>
    private sealed class MarginStyleContextKey : IEquatable<MarginStyleContextKey> {
        private readonly HtmlComputedStyleSet _styles;
        private readonly HtmlComputedStyle _root;
        private readonly HtmlCssPageRule[] _rules;
        private readonly int _hashCode;

        internal MarginStyleContextKey(HtmlComputedStyleSet styles, HtmlComputedStyle root, HtmlCssPageRule[] rules) {
            _styles = styles;
            _root = root;
            _rules = rules;
            unchecked {
                int hash = RuntimeHelpers.GetHashCode(styles) * 397 ^ RuntimeHelpers.GetHashCode(root);
                foreach (HtmlCssPageRule rule in rules) hash = hash * 397 ^ RuntimeHelpers.GetHashCode(rule);
                _hashCode = hash;
            }
        }

        public bool Equals(MarginStyleContextKey? other) {
            if (other == null || !ReferenceEquals(_styles, other._styles) || !ReferenceEquals(_root, other._root)
                || _rules.Length != other._rules.Length) return false;
            for (int index = 0; index < _rules.Length; index++) {
                if (!ReferenceEquals(_rules[index], other._rules[index])) return false;
            }
            return true;
        }

        public override bool Equals(object? obj) => obj is MarginStyleContextKey other && Equals(other);
        public override int GetHashCode() => _hashCode;
    }
}
