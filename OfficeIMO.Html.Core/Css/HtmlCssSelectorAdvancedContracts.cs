using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading;
using OfficeIMO.Html.Dom;

namespace OfficeIMO.Html.Css;

/// <summary>Immutable stylesheet-scoped namespace bindings for CSS selectors.</summary>
public sealed class HtmlCssNamespaceContext {
    private readonly IReadOnlyDictionary<string, string> _prefixes;

    /// <summary>Creates namespace bindings. A null default leaves unprefixed type selectors namespace-unconstrained.</summary>
    public HtmlCssNamespaceContext(string? defaultNamespaceUri = null, IReadOnlyDictionary<string, string>? prefixes = null) {
        DefaultNamespaceUri = defaultNamespaceUri;
        var copy = new Dictionary<string, string>(StringComparer.Ordinal);
        if (prefixes != null) {
            foreach (KeyValuePair<string, string> binding in prefixes) {
                if (string.IsNullOrEmpty(binding.Key)) throw new ArgumentException("Namespace prefixes cannot be empty.", nameof(prefixes));
                if (binding.Value == null) throw new ArgumentException("Namespace URIs cannot be null.", nameof(prefixes));
                copy[binding.Key] = binding.Value;
            }
        }
        _prefixes = new ReadOnlyDictionary<string, string>(copy);
    }

    /// <summary>The default namespace URI, or null when no default namespace is declared.</summary>
    public string? DefaultNamespaceUri { get; }
    /// <summary>Case-sensitive prefix bindings.</summary>
    public IReadOnlyDictionary<string, string> Prefixes => _prefixes;
    /// <summary>Resolves a case-sensitive namespace prefix.</summary>
    public bool TryGetNamespaceUri(string prefix, out string namespaceUri) => _prefixes.TryGetValue(prefix, out namespaceUri!);

    /// <summary>Reads valid top-level @namespace declarations from an owned stylesheet.</summary>
    public static HtmlCssNamespaceContext FromStyleSheet(HtmlCssStyleSheet styleSheet) {
        if (styleSheet == null) throw new ArgumentNullException(nameof(styleSheet));
        string? defaultNamespace = null;
        var prefixes = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (HtmlCssRule rule in styleSheet.Rules) {
            if (rule is not HtmlCssAtRule atRule) break;
            if (!string.Equals(atRule.Name, "namespace", StringComparison.OrdinalIgnoreCase)) {
                if (string.Equals(atRule.Name, "charset", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(atRule.Name, "import", StringComparison.OrdinalIgnoreCase)) continue;
                break;
            }
            if (TryParseDeclaration(atRule.PreludeText, out string? prefix, out string? uri)) {
                if (prefix == null) defaultNamespace = uri;
                else prefixes[prefix] = uri!;
            }
        }
        return new HtmlCssNamespaceContext(defaultNamespace, prefixes);
    }

    private static bool TryParseDeclaration(string prelude, out string? prefix, out string? uri) {
        prefix = null; uri = null;
        IReadOnlyList<HtmlCssToken> tokens = HtmlCssTokenizer.Tokenize(prelude);
        var significant = tokens.Where(token => token.Kind != HtmlCssTokenKind.Whitespace
            && token.Kind != HtmlCssTokenKind.Comment && token.Kind != HtmlCssTokenKind.EndOfFile).ToList();
        int index = 0;
        if (significant.Count > 0 && significant[0].Kind == HtmlCssTokenKind.Identifier) prefix = significant[index++].Value;
        if (index >= significant.Count) return false;
        HtmlCssToken value = significant[index++];
        if (value.Kind == HtmlCssTokenKind.String) uri = value.Value;
        else if (value.Kind == HtmlCssTokenKind.Url) uri = value.Value;
        else if (value.Kind == HtmlCssTokenKind.Function && string.Equals(value.Value, "url", StringComparison.OrdinalIgnoreCase)
            && index + 1 < significant.Count && significant[index].Kind == HtmlCssTokenKind.String
            && significant[index + 1].Kind == HtmlCssTokenKind.CloseParenthesis) {
            uri = significant[index].Value;
            index += 2;
        } else return false;
        return index == significant.Count && uri != null;
    }
}

internal enum HtmlCssNamespaceConstraintKind { Any, None, Exact }

internal readonly struct HtmlCssNamespaceConstraint {
    private HtmlCssNamespaceConstraint(HtmlCssNamespaceConstraintKind kind, string uri) { Kind = kind; Uri = uri; }
    internal static HtmlCssNamespaceConstraint Any { get; } = new HtmlCssNamespaceConstraint(HtmlCssNamespaceConstraintKind.Any, string.Empty);
    internal static HtmlCssNamespaceConstraint None { get; } = new HtmlCssNamespaceConstraint(HtmlCssNamespaceConstraintKind.None, string.Empty);
    internal static HtmlCssNamespaceConstraint Exact(string uri) => new HtmlCssNamespaceConstraint(HtmlCssNamespaceConstraintKind.Exact, uri);
    internal HtmlCssNamespaceConstraintKind Kind { get; }
    internal string Uri { get; }
    internal bool Matches(string uri) => Kind == HtmlCssNamespaceConstraintKind.Any
        || Kind == HtmlCssNamespaceConstraintKind.None && uri.Length == 0
        || Kind == HtmlCssNamespaceConstraintKind.Exact && string.Equals(Uri, uri, StringComparison.Ordinal);
    internal void AppendIdentity(System.Text.StringBuilder builder) => builder.Append('n').Append((int)Kind).Append(':').Append(Uri.Length).Append(':').Append(Uri).Append(';');
}

internal enum HtmlCssPseudoClassKind {
    Root, Empty, FirstChild, LastChild, OnlyChild,
    FirstOfType, LastOfType, OnlyOfType,
    NthChild, NthLastChild, NthOfType, NthLastOfType,
    Is, Where, Not, Lang, Provider
}

internal readonly struct HtmlCssAnPlusB {
    internal HtmlCssAnPlusB(int a, int b) { A = a; B = b; }
    internal int A { get; }
    internal int B { get; }
    internal bool Matches(int index) {
        if (A == 0) return index == B;
        long difference = (long)index - B;
        long coefficient = A;
        return difference % coefficient == 0 && difference / coefficient >= 0;
    }
    public override string ToString() => A + "n" + (B < 0 ? B.ToString() : "+" + B);
}

internal sealed class HtmlCssPseudoClassSelector {
    internal HtmlCssPseudoClassSelector(HtmlCssPseudoClassKind kind, HtmlCssAnPlusB? formula = null,
        IReadOnlyList<HtmlCssSelector>? selectors = null, string? stringArgument = null, string? providerSource = null) {
        Kind = kind; Formula = formula; Selectors = selectors ?? Array.Empty<HtmlCssSelector>();
        StringArgument = stringArgument; ProviderSource = providerSource;
    }
    internal HtmlCssPseudoClassKind Kind { get; }
    internal HtmlCssAnPlusB? Formula { get; }
    internal IReadOnlyList<HtmlCssSelector> Selectors { get; }
    internal string? StringArgument { get; }
    internal string? ProviderSource { get; }

    internal bool Matches(IHtmlCssSelectorElement element, HtmlCssSelectorMatchContext context) {
        context.ThrowIfCancellationRequested();
        HtmlCssSiblingPosition position;
        switch (Kind) {
            case HtmlCssPseudoClassKind.Root: return element.IsDocumentElement;
            case HtmlCssPseudoClassKind.Empty: return !element.HasElementOrTextChild(context.RecordEvaluationAction, context.CancellationToken);
            case HtmlCssPseudoClassKind.FirstChild:
                position = context.GetSiblingPosition(element); return position.Index == 1;
            case HtmlCssPseudoClassKind.LastChild:
                position = context.GetSiblingPosition(element); return position.Index == position.Count;
            case HtmlCssPseudoClassKind.OnlyChild:
                position = context.GetSiblingPosition(element); return position.Count == 1;
            case HtmlCssPseudoClassKind.FirstOfType:
                position = context.GetSiblingPosition(element); return position.TypeIndex == 1;
            case HtmlCssPseudoClassKind.LastOfType:
                position = context.GetSiblingPosition(element); return position.TypeIndex == position.TypeCount;
            case HtmlCssPseudoClassKind.OnlyOfType:
                position = context.GetSiblingPosition(element); return position.TypeCount == 1;
            case HtmlCssPseudoClassKind.NthChild:
                position = context.GetSiblingPosition(element); return Formula!.Value.Matches(position.Index);
            case HtmlCssPseudoClassKind.NthLastChild:
                position = context.GetSiblingPosition(element); return Formula!.Value.Matches(position.Count - position.Index + 1);
            case HtmlCssPseudoClassKind.NthOfType:
                position = context.GetSiblingPosition(element); return Formula!.Value.Matches(position.TypeIndex);
            case HtmlCssPseudoClassKind.NthLastOfType:
                position = context.GetSiblingPosition(element); return Formula!.Value.Matches(position.TypeCount - position.TypeIndex + 1);
            case HtmlCssPseudoClassKind.Not:
                foreach (HtmlCssSelector selector in Selectors)
                    if (selector.Matches(element, context)) return false;
                return true;
            case HtmlCssPseudoClassKind.Lang:
                return StringArgument != null && MatchesLanguage(element, StringArgument, context);
            case HtmlCssPseudoClassKind.Provider:
                return ProviderSource != null && context.MatchesProviderSelector(element, ProviderSource);
            default:
                foreach (HtmlCssSelector selector in Selectors)
                    if (selector.Matches(element, context)) return true;
                return false;
        }
    }

    internal void AppendIdentity(System.Text.StringBuilder builder, bool ignoreAttributeModifiers) {
        builder.Append('p').Append((int)Kind).Append(':');
        if (Formula.HasValue) builder.Append(Formula.Value.ToString());
        if (StringArgument != null) builder.Append(StringArgument.Length).Append(':').Append(StringArgument);
        if (ProviderSource != null) builder.Append(ProviderSource);
        foreach (HtmlCssSelector selector in Selectors)
            builder.Append('{').Append(ignoreAttributeModifiers ? selector.CompatibilityIdentity : selector.Identity).Append('}');
        builder.Append(';');
    }

    private static bool MatchesLanguage(IHtmlCssSelectorElement element, string range, HtmlCssSelectorMatchContext context) {
        const string XmlNamespace = "http://www.w3.org/XML/1998/namespace";
        for (IHtmlCssSelectorElement? current = element; current != null; current = current.ParentElement) {
            context.ThrowIfCancellationRequested();
            context.RecordEvaluation();
            string? language = null;
            bool declared = false;
            foreach (HtmlCssSelectorAttributeValue attribute in current.Attributes) {
                context.ThrowIfCancellationRequested();
                context.RecordEvaluation();
                if (attribute.NamespaceUri.Length == 0 && HtmlCssAscii.EqualsIgnoreCase(attribute.LocalName, "lang")) {
                    language = attribute.Value; declared = true; break;
                }
            }
            if (!declared) {
                foreach (HtmlCssSelectorAttributeValue attribute in current.Attributes) {
                    context.ThrowIfCancellationRequested();
                    context.RecordEvaluation();
                    if (attribute.NamespaceUri == XmlNamespace && attribute.LocalName == "lang") {
                        language = attribute.Value; declared = true; break;
                    }
                }
            }
            if (declared) return LanguageRangeMatches(language ?? string.Empty, range);
        }
        return false;
    }

    private static bool LanguageRangeMatches(string language, string range) =>
        HtmlCssAscii.EqualsIgnoreCase(language, range)
        || language.Length > range.Length && language[range.Length] == '-'
        && HtmlCssAscii.StartsWithIgnoreCase(language, range);

}
