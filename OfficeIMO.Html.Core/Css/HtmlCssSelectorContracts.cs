using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using OfficeIMO.Html.Dom;

namespace OfficeIMO.Html.Css;

/// <summary>Outcome of parsing one selector against the OfficeIMO-owned selector subset.</summary>
public enum HtmlCssSelectorParseStatus {
    /// <summary>The selector is fully represented by the owned selector model.</summary>
    Parsed,
    /// <summary>The selector is valid CSS syntax but uses a selector feature outside the owned subset.</summary>
    Unsupported,
    /// <summary>The selector is malformed.</summary>
    InvalidSyntax
}

/// <summary>Combinators supported by the first owned selector-matching slice.</summary>
public enum HtmlCssCombinator {
    /// <summary>A descendant of the element on the left.</summary>
    Descendant,
    /// <summary>A direct child of the element on the left.</summary>
    Child,
    /// <summary>The immediately following element sibling.</summary>
    NextSibling,
    /// <summary>A later element sibling.</summary>
    SubsequentSibling
}

/// <summary>CSS selector specificity as an id, class, and type tuple.</summary>
public readonly struct HtmlCssSelectorSpecificity : IEquatable<HtmlCssSelectorSpecificity> {
    internal HtmlCssSelectorSpecificity(int ids, int classes, int types) {
        Ids = ids;
        Classes = classes;
        Types = types;
    }

    /// <summary>Number of id selectors.</summary>
    public int Ids { get; }
    /// <summary>Number of class and attribute selectors.</summary>
    public int Classes { get; }
    /// <summary>Number of type selectors.</summary>
    public int Types { get; }
    /// <inheritdoc />
    public bool Equals(HtmlCssSelectorSpecificity other) => Ids == other.Ids && Classes == other.Classes && Types == other.Types;
    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is HtmlCssSelectorSpecificity other && Equals(other);
    /// <inheritdoc />
    public override int GetHashCode() => ((Ids * 397) ^ Classes) * 397 ^ Types;
    /// <inheritdoc />
    public override string ToString() => Ids + "," + Classes + "," + Types;

    internal static HtmlCssSelectorSpecificity Add(HtmlCssSelectorSpecificity left, HtmlCssSelectorSpecificity right) =>
        new HtmlCssSelectorSpecificity(checked(left.Ids + right.Ids), checked(left.Classes + right.Classes), checked(left.Types + right.Types));

    internal static HtmlCssSelectorSpecificity Max(HtmlCssSelectorSpecificity left, HtmlCssSelectorSpecificity right) {
        if (left.Ids != right.Ids) return left.Ids > right.Ids ? left : right;
        if (left.Classes != right.Classes) return left.Classes > right.Classes ? left : right;
        return left.Types >= right.Types ? left : right;
    }
}

/// <summary>A parsed selector that can match the owned DOM without a selector provider.</summary>
public sealed class HtmlCssSelector {
    private readonly IReadOnlyList<HtmlCssSelectorCompound> _compounds;
    private readonly IReadOnlyList<HtmlCssCombinator> _combinators;

    internal HtmlCssSelector(
        string source,
        IReadOnlyList<HtmlCssSelectorCompound> compounds,
        IReadOnlyList<HtmlCssCombinator> combinators,
        HtmlCssSelectorSpecificity specificity) {
        Source = source;
        _compounds = compounds;
        _combinators = combinators;
        Specificity = specificity;
        Identity = BuildIdentity(compounds, combinators);
        CompatibilityIdentity = BuildIdentity(compounds, combinators, ignoreAttributeModifiers: true);
    }

    /// <summary>The trimmed selector source.</summary>
    public string Source { get; }
    /// <summary>The selector specificity computed from the parsed selector model.</summary>
    public HtmlCssSelectorSpecificity Specificity { get; }
    internal string Identity { get; }
    internal string CompatibilityIdentity { get; }
    internal bool RequiresProviderMatching => _compounds.Any(compound => compound.RequiresProviderMatching);
    /// <summary>Number of compound selectors in the complex selector.</summary>
    public int CompoundCount => _compounds.Count;
    /// <summary>Combinators in left-to-right source order.</summary>
    public IReadOnlyList<HtmlCssCombinator> Combinators => _combinators;

    /// <summary>Matches an element using only owned DOM state.</summary>
    public bool Matches(HtmlElement element) => Matches(element, CancellationToken.None);

    /// <summary>Matches an element using only owned DOM state and observes cancellation while traversing the tree.</summary>
    public bool Matches(HtmlElement element, CancellationToken cancellationToken) {
        if (element == null) throw new ArgumentNullException(nameof(element));
        return Matches(new OwnedSelectorElement(element), null, cancellationToken);
    }

    internal bool Matches(
        IHtmlCssSelectorElement element,
        Action? recordEvaluation = null,
        CancellationToken cancellationToken = default) =>
        Matches(element, new HtmlCssSelectorMatchContext(recordEvaluation, cancellationToken));

    internal bool Matches(IHtmlCssSelectorElement element, HtmlCssSelectorMatchContext context) {
        if (_compounds.Count == 1) {
            context.ThrowIfCancellationRequested();
            context.RecordEvaluation();
            return _compounds[0].Matches(element, context);
        }
        return MatchAt(_compounds.Count - 1, element, context);
    }

    private bool MatchAt(
        int compoundIndex,
        IHtmlCssSelectorElement element,
        HtmlCssSelectorMatchContext context) {
        context.ThrowIfCancellationRequested();
        var state = new MatchState(compoundIndex, element.Identity);
        if (context.TryGetMatchResult(this, state, out bool retained)) return retained;
        context.RecordEvaluation();
        bool result;
        if (!_compounds[compoundIndex].Matches(element, context)) result = false;
        else if (compoundIndex == 0) result = true;
        else {
        switch (_combinators[compoundIndex - 1]) {
            case HtmlCssCombinator.Child:
                result = element.ParentElement != null && MatchAt(compoundIndex - 1, element.ParentElement, context);
                break;
            case HtmlCssCombinator.NextSibling:
                IHtmlCssSelectorElement? previous = context.GetPreviousElementSibling(element);
                result = previous != null && MatchAt(compoundIndex - 1, previous, context);
                break;
            case HtmlCssCombinator.SubsequentSibling:
                result = false;
                for (IHtmlCssSelectorElement? sibling = context.GetPreviousElementSibling(element); sibling != null; sibling = context.GetPreviousElementSibling(sibling))
                    if (MatchAt(compoundIndex - 1, sibling, context)) { result = true; break; }
                break;
            default:
                result = false;
                for (IHtmlCssSelectorElement? ancestor = element.ParentElement; ancestor != null; ancestor = ancestor.ParentElement)
                    if (MatchAt(compoundIndex - 1, ancestor, context)) { result = true; break; }
                break;
        }
        }
        context.SetMatchResult(this, state, result);
        return result;
    }

    internal readonly struct MatchState : IEquatable<MatchState> {
        internal MatchState(int compoundIndex, object element) { CompoundIndex = compoundIndex; Element = element; }
        private int CompoundIndex { get; }
        private object Element { get; }
        public bool Equals(MatchState other) => CompoundIndex == other.CompoundIndex && ReferenceEquals(Element, other.Element);
        public override bool Equals(object? obj) => obj is MatchState other && Equals(other);
        public override int GetHashCode() => (CompoundIndex * 397) ^ RuntimeHelpers.GetHashCode(Element);
    }

    private static string BuildIdentity(
        IReadOnlyList<HtmlCssSelectorCompound> compounds,
        IReadOnlyList<HtmlCssCombinator> combinators,
        bool ignoreAttributeModifiers = false) {
        var builder = new System.Text.StringBuilder();
        for (int index = 0; index < compounds.Count; index++) {
            if (index > 0) builder.Append('/').Append((int)combinators[index - 1]).Append('/');
            compounds[index].AppendIdentity(builder, ignoreAttributeModifiers);
        }
        return builder.ToString();
    }
}

/// <summary>A parsed comma-separated selector list whose members can match the owned DOM.</summary>
public sealed class HtmlCssSelectorList {
    private readonly IReadOnlyList<HtmlCssSelector> _selectors;

    internal HtmlCssSelectorList(string source, IReadOnlyList<HtmlCssSelector> selectors) {
        Source = source;
        _selectors = selectors;
    }

    /// <summary>The trimmed selector-list source.</summary>
    public string Source { get; }
    /// <summary>Selectors in authored order.</summary>
    public IReadOnlyList<HtmlCssSelector> Selectors => _selectors;
    /// <summary>Matches when any selector in the list matches the element.</summary>
    public bool Matches(HtmlElement element) => Matches(element, CancellationToken.None);
    /// <summary>Matches when any selector in the list matches, observing cancellation while traversing the tree.</summary>
    public bool Matches(HtmlElement element, CancellationToken cancellationToken) {
        if (element == null) throw new ArgumentNullException(nameof(element));
        var adapted = new OwnedSelectorElement(element);
        var context = new HtmlCssSelectorMatchContext(null, cancellationToken);
        foreach (HtmlCssSelector selector in _selectors) {
            cancellationToken.ThrowIfCancellationRequested();
            if (selector.Matches(adapted, context)) return true;
        }
        return false;
    }

    internal string Identity => string.Join(",", _selectors.Select(selector => selector.Identity));
    internal string CompatibilityIdentity => string.Join(",", _selectors.Select(selector => selector.CompatibilityIdentity));
}

/// <summary>Result of parsing one selector without invoking a selector provider.</summary>
public sealed class HtmlCssSelectorParseResult {
    internal HtmlCssSelectorParseResult(HtmlCssSelectorParseStatus status, string source, HtmlCssSelector? selector, string? reason) {
        Status = status;
        Source = source;
        Selector = selector;
        Reason = reason;
    }

    /// <summary>Parse outcome.</summary>
    public HtmlCssSelectorParseStatus Status { get; }
    /// <summary>Trimmed authored selector.</summary>
    public string Source { get; }
    /// <summary>Parsed selector when <see cref="Status"/> is <see cref="HtmlCssSelectorParseStatus.Parsed"/>.</summary>
    public HtmlCssSelector? Selector { get; }
    /// <summary>Stable short reason for an unsupported or malformed selector.</summary>
    public string? Reason { get; }
    /// <summary>Whether matching can run through the owned selector engine.</summary>
    public bool IsSupported => Status == HtmlCssSelectorParseStatus.Parsed;
}

/// <summary>Result of parsing a selector list without invoking a selector provider.</summary>
public sealed class HtmlCssSelectorListParseResult {
    internal HtmlCssSelectorListParseResult(HtmlCssSelectorParseStatus status, string source, HtmlCssSelectorList? selectorList, string? reason) {
        Status = status;
        Source = source;
        SelectorList = selectorList;
        Reason = reason;
    }

    /// <summary>Parse outcome for the complete list.</summary>
    public HtmlCssSelectorParseStatus Status { get; }
    /// <summary>Trimmed authored selector-list source.</summary>
    public string Source { get; }
    /// <summary>Parsed selector list when every member is represented by the owned selector model.</summary>
    public HtmlCssSelectorList? SelectorList { get; }
    /// <summary>Stable short reason for an unsupported or malformed selector list.</summary>
    public string? Reason { get; }
    /// <summary>Whether matching can run through the owned selector engine.</summary>
    public bool IsSupported => Status == HtmlCssSelectorParseStatus.Parsed;
}

/// <summary>Resource bounds for parsing a standalone selector.</summary>
public sealed class HtmlCssSelectorOptions {
    /// <summary>Maximum UTF-16 selector length.</summary>
    public int MaxInputCharacters { get; set; } = 64 * 1024;
    /// <summary>Maximum lexical tokens, including trivia.</summary>
    public int MaxTokens { get; set; } = 4096;
    /// <summary>Maximum compound selectors in one complex selector.</summary>
    public int MaxCompounds { get; set; } = 128;
    /// <summary>Maximum simple selectors across all compounds.</summary>
    public int MaxSimpleSelectors { get; set; } = 512;
    /// <summary>Maximum selectors in one comma-separated selector list, including nested logical lists.</summary>
    public int MaxSelectors { get; set; } = 256;
    /// <summary>Maximum nesting depth of selector-taking pseudo-classes.</summary>
    public int MaxNestingDepth { get; set; } = 16;
    /// <summary>Stylesheet-scoped namespace bindings used by qualified type and attribute selectors.</summary>
    public HtmlCssNamespaceContext? Namespaces { get; set; }

    internal HtmlCssSelectorOptions Clone() => new HtmlCssSelectorOptions {
        MaxInputCharacters = MaxInputCharacters,
        MaxTokens = MaxTokens,
        MaxCompounds = MaxCompounds,
        MaxSimpleSelectors = MaxSimpleSelectors,
        MaxSelectors = MaxSelectors,
        MaxNestingDepth = MaxNestingDepth,
        Namespaces = Namespaces
    };

    internal void Validate() {
        if (MaxInputCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(MaxInputCharacters));
        if (MaxTokens <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTokens));
        if (MaxCompounds <= 0) throw new ArgumentOutOfRangeException(nameof(MaxCompounds));
        if (MaxSimpleSelectors <= 0) throw new ArgumentOutOfRangeException(nameof(MaxSimpleSelectors));
        if (MaxSelectors <= 0) throw new ArgumentOutOfRangeException(nameof(MaxSelectors));
        if (MaxNestingDepth <= 0) throw new ArgumentOutOfRangeException(nameof(MaxNestingDepth));
    }
}

/// <summary>A selector exceeded a caller-specified structural bound.</summary>
public sealed class HtmlCssSelectorLimitException : InvalidOperationException {
    /// <summary>Creates a selector limit failure.</summary>
    public HtmlCssSelectorLimitException(string limitName, long actual, long maximum)
        : base($"CSS selector {limitName} limit exceeded ({actual} > {maximum}).") {
        LimitName = limitName;
        Actual = actual;
        Maximum = maximum;
    }
    /// <summary>The exhausted option.</summary>
    public string LimitName { get; }
    /// <summary>The observed amount of work.</summary>
    public long Actual { get; }
    /// <summary>The configured maximum.</summary>
    public long Maximum { get; }
}

/// <summary>
/// Shares cancellable, accounted DOM traversal results across selector matches in one operation.
/// </summary>
internal sealed class HtmlCssSelectorMatchContext {
    private const int MaximumRetainedMatchStates = 262_144;
    private readonly Action? _recordEvaluation;
    private readonly CancellationToken _cancellationToken;
    private readonly Func<IHtmlCssSelectorElement, string, bool>? _providerMatcher;
    private readonly Dictionary<object, SiblingSet> _siblings =
        new Dictionary<object, SiblingSet>(ReferenceIdentityComparer.Instance);
    private readonly Dictionary<HtmlCssSelector, Dictionary<HtmlCssSelector.MatchState, bool>> _matchResults =
        new Dictionary<HtmlCssSelector, Dictionary<HtmlCssSelector.MatchState, bool>>();
    private int _retainedMatchStates;

    internal HtmlCssSelectorMatchContext(Action? recordEvaluation, CancellationToken cancellationToken,
        Func<IHtmlCssSelectorElement, string, bool>? providerMatcher = null) {
        _recordEvaluation = recordEvaluation;
        _cancellationToken = cancellationToken;
        _providerMatcher = providerMatcher;
    }

    internal Action? RecordEvaluationAction => _recordEvaluation;
    internal CancellationToken CancellationToken => _cancellationToken;
    internal void RecordEvaluation() => _recordEvaluation?.Invoke();
    internal void ThrowIfCancellationRequested() => _cancellationToken.ThrowIfCancellationRequested();
    internal bool TryGetMatchResult(HtmlCssSelector selector, HtmlCssSelector.MatchState state, out bool result) {
        if (_matchResults.TryGetValue(selector, out Dictionary<HtmlCssSelector.MatchState, bool>? matches))
            return matches.TryGetValue(state, out result);
        result = false;
        return false;
    }

    internal void SetMatchResult(HtmlCssSelector selector, HtmlCssSelector.MatchState state, bool result) {
        // A context belongs to one selector operation over a stable tree. Retaining a bounded
        // number of ancestor results avoids repeated walks across sibling elements.
        if (_retainedMatchStates >= MaximumRetainedMatchStates) {
            _matchResults.Clear();
            _retainedMatchStates = 0;
        }
        if (!_matchResults.TryGetValue(selector, out Dictionary<HtmlCssSelector.MatchState, bool>? matches)) {
            matches = new Dictionary<HtmlCssSelector.MatchState, bool>();
            _matchResults.Add(selector, matches);
        }
        if (!matches.ContainsKey(state)) {
            matches.Add(state, result);
            _retainedMatchStates++;
        }
    }
    internal bool MatchesProviderSelector(IHtmlCssSelectorElement element, string selector) {
        ThrowIfCancellationRequested();
        RecordEvaluation();
        return _providerMatcher != null && _providerMatcher(element, selector);
    }

    internal IHtmlCssSelectorElement? GetPreviousElementSibling(IHtmlCssSelectorElement element) {
        SiblingPosition? position = GetPosition(element);
        return position.HasValue && position.Value.Index > 0
            ? position.Value.Set.Elements[position.Value.Index - 1]
            : null;
    }

    internal IHtmlCssSelectorElement? GetNextElementSibling(IHtmlCssSelectorElement element) {
        SiblingPosition? position = GetPosition(element);
        return position.HasValue && position.Value.Index + 1 < position.Value.Set.Elements.Count
            ? position.Value.Set.Elements[position.Value.Index + 1]
            : null;
    }

    internal HtmlCssSiblingPosition GetSiblingPosition(IHtmlCssSelectorElement element) {
        SiblingPosition? position = GetPosition(element);
        if (!position.HasValue) return new HtmlCssSiblingPosition(1, 1, 1, 1);
        SiblingInfo info = position.Value.Set.Positions[element.Identity];
        return new HtmlCssSiblingPosition(
            position.Value.Index + 1,
            position.Value.Set.Elements.Count,
            info.TypeIndex,
            position.Value.Set.TypeCounts[info.TypeKey]);
    }

    private SiblingPosition? GetPosition(IHtmlCssSelectorElement element) {
        ThrowIfCancellationRequested();
        IHtmlCssSelectorElement? parent = element.ParentElement;
        if (parent == null) return null;
        if (!_siblings.TryGetValue(parent.Identity, out SiblingSet? set)) {
            IReadOnlyList<IHtmlCssSelectorElement> elements = parent.GetElementChildren(_recordEvaluation, _cancellationToken);
            set = new SiblingSet(elements);
            _siblings.Add(parent.Identity, set);
        }
        return set.Positions.TryGetValue(element.Identity, out SiblingInfo info)
            ? new SiblingPosition(set, info.Index)
            : null;
    }

    private sealed class SiblingSet {
        internal SiblingSet(IReadOnlyList<IHtmlCssSelectorElement> elements) {
            Elements = elements;
            Positions = new Dictionary<object, SiblingInfo>(ReferenceIdentityComparer.Instance);
            TypeCounts = new Dictionary<string, int>(StringComparer.Ordinal);
            for (int index = 0; index < elements.Count; index++) {
                IHtmlCssSelectorElement element = elements[index];
                string typeKey = BuildTypeKey(element);
                TypeCounts.TryGetValue(typeKey, out int typeIndex);
                typeIndex++;
                TypeCounts[typeKey] = typeIndex;
                Positions[element.Identity] = new SiblingInfo(index, typeIndex, typeKey);
            }
        }
        internal IReadOnlyList<IHtmlCssSelectorElement> Elements { get; }
        internal Dictionary<object, SiblingInfo> Positions { get; }
        internal Dictionary<string, int> TypeCounts { get; }
    }

    private readonly struct SiblingPosition {
        internal SiblingPosition(SiblingSet set, int index) { Set = set; Index = index; }
        internal SiblingSet Set { get; }
        internal int Index { get; }
    }

    private readonly struct SiblingInfo {
        internal SiblingInfo(int index, int typeIndex, string typeKey) { Index = index; TypeIndex = typeIndex; TypeKey = typeKey; }
        internal int Index { get; }
        internal int TypeIndex { get; }
        internal string TypeKey { get; }
    }

    private static string BuildTypeKey(IHtmlCssSelectorElement element) => element.NamespaceUri + "\u001f"
        + (element.NamespaceUri == HtmlElement.HtmlNamespace ? HtmlCssAscii.ToLowerInvariant(element.LocalName) : element.LocalName);

    private sealed class ReferenceIdentityComparer : IEqualityComparer<object> {
        internal static ReferenceIdentityComparer Instance { get; } = new ReferenceIdentityComparer();
        public new bool Equals(object? left, object? right) => ReferenceEquals(left, right);
        public int GetHashCode(object value) => RuntimeHelpers.GetHashCode(value);
    }
}

internal readonly struct HtmlCssSiblingPosition {
    internal HtmlCssSiblingPosition(int index, int count, int typeIndex, int typeCount) {
        Index = index; Count = count; TypeIndex = typeIndex; TypeCount = typeCount;
    }
    internal int Index { get; }
    internal int Count { get; }
    internal int TypeIndex { get; }
    internal int TypeCount { get; }
}

internal interface IHtmlCssSelectorElement {
    object Identity { get; }
    string LocalName { get; }
    string NamespaceUri { get; }
    string Id { get; }
    IHtmlCssSelectorElement? ParentElement { get; }
    IReadOnlyList<IHtmlCssSelectorElement> GetElementChildren(Action? recordEvaluation, CancellationToken cancellationToken);
    bool IsDocumentElement { get; }
    bool HasElementOrTextChild(Action? recordEvaluation, CancellationToken cancellationToken);
    bool HasClass(string name);
    IReadOnlyList<HtmlCssSelectorAttributeValue> Attributes { get; }
}

internal readonly struct HtmlCssSelectorAttributeValue {
    internal HtmlCssSelectorAttributeValue(string localName, string namespaceUri, string value) {
        LocalName = localName; NamespaceUri = namespaceUri; Value = value;
    }
    internal string LocalName { get; }
    internal string NamespaceUri { get; }
    internal string Value { get; }
}

internal sealed class OwnedSelectorElement : IHtmlCssSelectorElement {
    private readonly HtmlElement _element;
    private IReadOnlyList<HtmlCssSelectorAttributeValue>? _attributes;
    internal OwnedSelectorElement(HtmlElement element) { _element = element; }
    public object Identity => _element;
    public string LocalName => _element.LocalName;
    public string NamespaceUri => _element.NamespaceUri;
    public string Id => _element.Id;
    public IHtmlCssSelectorElement? ParentElement => _element.ParentElement == null ? null : new OwnedSelectorElement(_element.ParentElement);
    public IReadOnlyList<IHtmlCssSelectorElement> GetElementChildren(Action? recordEvaluation, CancellationToken cancellationToken) {
        var children = new List<IHtmlCssSelectorElement>();
        foreach (HtmlNode child in _element.ChildNodes) {
            cancellationToken.ThrowIfCancellationRequested();
            recordEvaluation?.Invoke();
            if (child is HtmlElement element) children.Add(new OwnedSelectorElement(element));
        }
        return children;
    }
    public bool IsDocumentElement => ReferenceEquals(_element.Document.DocumentElement, _element);
    public bool HasElementOrTextChild(Action? recordEvaluation, CancellationToken cancellationToken) {
        foreach (HtmlNode child in _element.ChildNodes) {
            cancellationToken.ThrowIfCancellationRequested();
            recordEvaluation?.Invoke();
            if (child is HtmlElement) return true;
            if (child.Kind == HtmlNodeKind.Text && child.TextContent.Length != 0) return true;
        }
        return false;
    }
    public IReadOnlyList<HtmlCssSelectorAttributeValue> Attributes => _attributes ??= _element.Attributes
        .Select(attribute => new HtmlCssSelectorAttributeValue(attribute.LocalName, attribute.NamespaceUri, attribute.Value)).ToArray();
    public bool HasClass(string name) {
        foreach (string item in _element.ClassList)
            if (string.Equals(item, name, StringComparison.Ordinal)) return true;
        return false;
    }
}

internal sealed class HtmlCssSelectorCompound {
    internal string? TypeName { get; set; }
    internal bool Universal { get; set; }
    internal HtmlCssNamespaceConstraint Namespace { get; set; } = HtmlCssNamespaceConstraint.Any;
    internal bool HasImplicitDefaultNamespace { get; set; }
    internal List<string> Ids { get; } = new List<string>();
    internal List<string> Classes { get; } = new List<string>();
    internal List<HtmlCssAttributeSelector> Attributes { get; } = new List<HtmlCssAttributeSelector>();
    internal List<HtmlCssPseudoClassSelector> PseudoClasses { get; } = new List<HtmlCssPseudoClassSelector>();
    internal bool RequiresProviderMatching => PseudoClasses.Any(pseudo => pseudo.Kind == HtmlCssPseudoClassKind.Provider
        || pseudo.Selectors.Any(selector => selector.RequiresProviderMatching));

    internal bool Matches(IHtmlCssSelectorElement element, HtmlCssSelectorMatchContext context) {
        if (!Namespace.Matches(element.NamespaceUri)) return false;
        if (TypeName != null) {
            if (element.NamespaceUri == HtmlElement.HtmlNamespace
                ? !HtmlCssAscii.EqualsIgnoreCase(element.LocalName, TypeName)
                : !string.Equals(element.LocalName, TypeName, StringComparison.Ordinal)) return false;
        }
        foreach (string id in Ids) if (!string.Equals(element.Id, id, StringComparison.Ordinal)) return false;
        foreach (string className in Classes) if (!element.HasClass(className)) return false;
        foreach (HtmlCssAttributeSelector attribute in Attributes)
            if (!attribute.Matches(element, context.RecordEvaluation, context.CancellationToken)) return false;
        foreach (HtmlCssPseudoClassSelector pseudoClass in PseudoClasses)
            if (!pseudoClass.Matches(element, context)) return false;
        return true;
    }

    internal void AppendIdentity(System.Text.StringBuilder builder, bool ignoreAttributeModifiers) {
        Append(builder, "t", TypeName);
        if (Universal) builder.Append("u;");
        Namespace.AppendIdentity(builder);
        foreach (string id in Ids) Append(builder, "i", id);
        foreach (string className in Classes) Append(builder, "c", className);
        foreach (HtmlCssAttributeSelector attribute in Attributes) attribute.AppendIdentity(builder, ignoreAttributeModifiers);
        foreach (HtmlCssPseudoClassSelector pseudoClass in PseudoClasses) pseudoClass.AppendIdentity(builder, ignoreAttributeModifiers);
    }

    private static void Append(System.Text.StringBuilder builder, string kind, string? value) {
        if (value == null) return;
        builder.Append(kind).Append(value.Length).Append(':').Append(value).Append(';');
    }
}

internal enum HtmlCssAttributeOperator { Present, Equals, Includes, DashMatch, Prefix, Suffix, Substring }

internal enum HtmlCssAttributeCase { Default, Sensitive, Insensitive }

internal sealed class HtmlCssAttributeSelector {
    private static readonly HashSet<string> HtmlAsciiInsensitiveValues = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "accept", "accept-charset", "align", "alink", "axis", "bgcolor", "charset", "checked", "clear",
        "codetype", "color", "compact", "declare", "defer", "dir", "direction", "disabled", "enctype",
        "face", "frame", "hreflang", "http-equiv", "lang", "language", "link", "media", "method", "multiple",
        "nohref", "noresize", "noshade", "nowrap", "readonly", "rel", "rev", "rules", "scope", "scrolling",
        "selected", "shape", "target", "text", "type", "valign", "valuetype", "vlink"
    };

    internal HtmlCssAttributeSelector(string name, HtmlCssNamespaceConstraint ns, HtmlCssAttributeOperator operation, string? value, HtmlCssAttributeCase caseMode) {
        Name = name; Namespace = ns; Operation = operation; Value = value; CaseMode = caseMode;
    }
    internal string Name { get; }
    internal HtmlCssNamespaceConstraint Namespace { get; }
    internal HtmlCssAttributeOperator Operation { get; }
    internal string? Value { get; }
    internal HtmlCssAttributeCase CaseMode { get; }

    internal bool Matches(IHtmlCssSelectorElement element, Action? recordEvaluation, CancellationToken cancellationToken) {
        foreach (HtmlCssSelectorAttributeValue attribute in element.Attributes) {
            cancellationToken.ThrowIfCancellationRequested();
            recordEvaluation?.Invoke();
            if (!Namespace.Matches(attribute.NamespaceUri)) continue;
            bool htmlNoNamespace = element.NamespaceUri == HtmlElement.HtmlNamespace && attribute.NamespaceUri.Length == 0;
            if (htmlNoNamespace ? !HtmlCssAscii.EqualsIgnoreCase(attribute.LocalName, Name)
                : !string.Equals(attribute.LocalName, Name, StringComparison.Ordinal)) continue;
            if (Operation == HtmlCssAttributeOperator.Present) return true;
            if (MatchesValue(attribute.Value, htmlNoNamespace)) return true;
        }
        return false;
    }

    private bool MatchesValue(string actual, bool htmlNoNamespace) {
        string expected = Value ?? string.Empty;
        bool insensitive = CaseMode == HtmlCssAttributeCase.Insensitive
            || CaseMode == HtmlCssAttributeCase.Default && htmlNoNamespace && HtmlAsciiInsensitiveValues.Contains(Name);
        if (expected.Length == 0 && (Operation == HtmlCssAttributeOperator.Prefix
            || Operation == HtmlCssAttributeOperator.Suffix || Operation == HtmlCssAttributeOperator.Substring)) return false;
        switch (Operation) {
            case HtmlCssAttributeOperator.Equals: return insensitive
                ? HtmlCssAscii.EqualsIgnoreCase(actual, expected) : string.Equals(actual, expected, StringComparison.Ordinal);
            case HtmlCssAttributeOperator.Includes:
                foreach (string token in actual.Split(new[] { ' ', '\t', '\r', '\n', '\f' }, StringSplitOptions.RemoveEmptyEntries))
                    if (insensitive ? HtmlCssAscii.EqualsIgnoreCase(token, expected) : string.Equals(token, expected, StringComparison.Ordinal)) return true;
                return false;
            case HtmlCssAttributeOperator.DashMatch:
                return (insensitive ? HtmlCssAscii.EqualsIgnoreCase(actual, expected) : string.Equals(actual, expected, StringComparison.Ordinal))
                    || actual.Length > expected.Length && actual[expected.Length] == '-'
                    && (insensitive ? HtmlCssAscii.StartsWithIgnoreCase(actual, expected) : actual.StartsWith(expected, StringComparison.Ordinal));
            case HtmlCssAttributeOperator.Prefix: return insensitive
                ? HtmlCssAscii.StartsWithIgnoreCase(actual, expected) : actual.StartsWith(expected, StringComparison.Ordinal);
            case HtmlCssAttributeOperator.Suffix: return insensitive
                ? HtmlCssAscii.EndsWithIgnoreCase(actual, expected) : actual.EndsWith(expected, StringComparison.Ordinal);
            default: return insensitive
                ? HtmlCssAscii.IndexOfIgnoreCase(actual, expected) >= 0 : actual.IndexOf(expected, StringComparison.Ordinal) >= 0;
        }
    }

    internal void AppendIdentity(System.Text.StringBuilder builder, bool ignoreCaseModifier) {
        builder.Append('a');
        Namespace.AppendIdentity(builder);
        builder.Append(Name.Length).Append(':').Append(Name).Append(':').Append((int)Operation)
            .Append(':').Append(ignoreCaseModifier ? 'x' : CaseMode == HtmlCssAttributeCase.Insensitive ? 'i'
                : CaseMode == HtmlCssAttributeCase.Sensitive ? 's' : 'd').Append(':');
        string value = Value ?? string.Empty;
        builder.Append(value.Length).Append(':').Append(value).Append(';');
    }
}

internal static class HtmlCssAscii {
    internal static string ToLowerInvariant(string value) {
        char[]? characters = null;
        for (int index = 0; index < value.Length; index++) {
            char folded = Fold(value[index]);
            if (folded == value[index]) continue;
            characters ??= value.ToCharArray();
            characters[index] = folded;
        }
        return characters == null ? value : new string(characters);
    }
    internal static bool EqualsIgnoreCase(string left, string right) => left.Length == right.Length && IndexOfIgnoreCase(left, right) == 0;
    internal static bool StartsWithIgnoreCase(string value, string prefix) =>
        prefix.Length <= value.Length && RegionEquals(value, 0, prefix);
    internal static bool EndsWithIgnoreCase(string value, string suffix) =>
        suffix.Length <= value.Length && RegionEquals(value, value.Length - suffix.Length, suffix);
    internal static int IndexOfIgnoreCase(string value, string search) {
        if (search.Length == 0) return 0;
        for (int start = 0; start <= value.Length - search.Length; start++)
            if (RegionEquals(value, start, search)) return start;
        return -1;
    }
    private static bool RegionEquals(string value, int start, string search) {
        for (int index = 0; index < search.Length; index++)
            if (Fold(value[start + index]) != Fold(search[index])) return false;
        return true;
    }
    private static char Fold(char value) => value is >= 'A' and <= 'Z' ? (char)(value + ('a' - 'A')) : value;
}
