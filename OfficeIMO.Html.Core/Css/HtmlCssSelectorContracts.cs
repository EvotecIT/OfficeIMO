using System;
using System.Collections.Generic;
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
        MatchAt(_compounds.Count - 1, element, new Dictionary<MatchState, bool>(), recordEvaluation, cancellationToken);

    private bool MatchAt(
        int compoundIndex,
        IHtmlCssSelectorElement element,
        IDictionary<MatchState, bool> results,
        Action? recordEvaluation,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var state = new MatchState(compoundIndex, element.Identity);
        if (results.TryGetValue(state, out bool retained)) return retained;
        recordEvaluation?.Invoke();
        bool result;
        if (!_compounds[compoundIndex].Matches(element)) result = false;
        else if (compoundIndex == 0) result = true;
        else {
        switch (_combinators[compoundIndex - 1]) {
            case HtmlCssCombinator.Child:
                result = element.ParentElement != null && MatchAt(compoundIndex - 1, element.ParentElement, results, recordEvaluation, cancellationToken);
                break;
            case HtmlCssCombinator.NextSibling:
                result = element.PreviousElementSibling != null && MatchAt(compoundIndex - 1, element.PreviousElementSibling, results, recordEvaluation, cancellationToken);
                break;
            case HtmlCssCombinator.SubsequentSibling:
                result = false;
                for (IHtmlCssSelectorElement? sibling = element.PreviousElementSibling; sibling != null; sibling = sibling.PreviousElementSibling)
                    if (MatchAt(compoundIndex - 1, sibling, results, recordEvaluation, cancellationToken)) { result = true; break; }
                break;
            default:
                result = false;
                for (IHtmlCssSelectorElement? ancestor = element.ParentElement; ancestor != null; ancestor = ancestor.ParentElement)
                    if (MatchAt(compoundIndex - 1, ancestor, results, recordEvaluation, cancellationToken)) { result = true; break; }
                break;
        }
        }
        results[state] = result;
        return result;
    }

    private readonly struct MatchState : IEquatable<MatchState> {
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

    internal HtmlCssSelectorOptions Clone() => new HtmlCssSelectorOptions {
        MaxInputCharacters = MaxInputCharacters,
        MaxTokens = MaxTokens,
        MaxCompounds = MaxCompounds,
        MaxSimpleSelectors = MaxSimpleSelectors
    };

    internal void Validate() {
        if (MaxInputCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(MaxInputCharacters));
        if (MaxTokens <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTokens));
        if (MaxCompounds <= 0) throw new ArgumentOutOfRangeException(nameof(MaxCompounds));
        if (MaxSimpleSelectors <= 0) throw new ArgumentOutOfRangeException(nameof(MaxSimpleSelectors));
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

internal interface IHtmlCssSelectorElement {
    object Identity { get; }
    string LocalName { get; }
    string NamespaceUri { get; }
    string Id { get; }
    IHtmlCssSelectorElement? ParentElement { get; }
    IHtmlCssSelectorElement? PreviousElementSibling { get; }
    bool HasClass(string name);
    bool HasAttributeInNoNamespace(string name);
    string? GetAttributeInNoNamespace(string name);
}

internal sealed class OwnedSelectorElement : IHtmlCssSelectorElement {
    private readonly HtmlElement _element;
    internal OwnedSelectorElement(HtmlElement element) { _element = element; }
    public object Identity => _element;
    public string LocalName => _element.LocalName;
    public string NamespaceUri => _element.NamespaceUri;
    public string Id => _element.Id;
    public IHtmlCssSelectorElement? ParentElement => _element.ParentElement == null ? null : new OwnedSelectorElement(_element.ParentElement);
    public IHtmlCssSelectorElement? PreviousElementSibling {
        get {
            for (HtmlNode? node = _element.PreviousSibling; node != null; node = node.PreviousSibling)
                if (node is HtmlElement element) return new OwnedSelectorElement(element);
            return null;
        }
    }
    public bool HasAttributeInNoNamespace(string name) => GetAttributeInNoNamespace(name) != null;
    public string? GetAttributeInNoNamespace(string name) {
        foreach (HtmlAttribute attribute in _element.Attributes) {
            if (attribute.NamespaceUri.Length != 0) continue;
            if (_element.NamespaceUri == HtmlElement.HtmlNamespace
                ? HtmlCssAscii.EqualsIgnoreCase(attribute.LocalName, name)
                : string.Equals(attribute.LocalName, name, StringComparison.Ordinal)) return attribute.Value;
        }
        return null;
    }
    public bool HasClass(string name) {
        foreach (string item in _element.ClassList)
            if (string.Equals(item, name, StringComparison.Ordinal)) return true;
        return false;
    }
}

internal sealed class HtmlCssSelectorCompound {
    internal string? TypeName { get; set; }
    internal bool Universal { get; set; }
    internal List<string> Ids { get; } = new List<string>();
    internal List<string> Classes { get; } = new List<string>();
    internal List<HtmlCssAttributeSelector> Attributes { get; } = new List<HtmlCssAttributeSelector>();

    internal bool Matches(IHtmlCssSelectorElement element) {
        if (TypeName != null) {
            if (element.NamespaceUri == HtmlElement.HtmlNamespace
                ? !HtmlCssAscii.EqualsIgnoreCase(element.LocalName, TypeName)
                : !string.Equals(element.LocalName, TypeName, StringComparison.Ordinal)) return false;
        }
        foreach (string id in Ids) if (!string.Equals(element.Id, id, StringComparison.Ordinal)) return false;
        foreach (string className in Classes) if (!element.HasClass(className)) return false;
        foreach (HtmlCssAttributeSelector attribute in Attributes)
            if (!attribute.Matches(element)) return false;
        return true;
    }

    internal void AppendIdentity(System.Text.StringBuilder builder, bool ignoreAttributeModifiers) {
        Append(builder, "t", TypeName);
        if (Universal) builder.Append("u;");
        foreach (string id in Ids) Append(builder, "i", id);
        foreach (string className in Classes) Append(builder, "c", className);
        foreach (HtmlCssAttributeSelector attribute in Attributes) attribute.AppendIdentity(builder, ignoreAttributeModifiers);
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

    internal HtmlCssAttributeSelector(string name, HtmlCssAttributeOperator operation, string? value, HtmlCssAttributeCase caseMode) {
        Name = name; Operation = operation; Value = value; CaseMode = caseMode;
    }
    internal string Name { get; }
    internal HtmlCssAttributeOperator Operation { get; }
    internal string? Value { get; }
    internal HtmlCssAttributeCase CaseMode { get; }

    internal bool Matches(IHtmlCssSelectorElement element) {
        if (!element.HasAttributeInNoNamespace(Name)) return false;
        if (Operation == HtmlCssAttributeOperator.Present) return true;
        string actual = element.GetAttributeInNoNamespace(Name) ?? string.Empty;
        string expected = Value ?? string.Empty;
        bool insensitive = CaseMode == HtmlCssAttributeCase.Insensitive
            || CaseMode == HtmlCssAttributeCase.Default && element.NamespaceUri == HtmlElement.HtmlNamespace
            && HtmlAsciiInsensitiveValues.Contains(Name);
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
        builder.Append('a').Append(Name.Length).Append(':').Append(Name).Append(':').Append((int)Operation)
            .Append(':').Append(ignoreCaseModifier ? 'x' : CaseMode == HtmlCssAttributeCase.Insensitive ? 'i'
                : CaseMode == HtmlCssAttributeCase.Sensitive ? 's' : 'd').Append(':');
        string value = Value ?? string.Empty;
        builder.Append(value.Length).Append(':').Append(value).Append(';');
    }
}

internal static class HtmlCssAscii {
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
