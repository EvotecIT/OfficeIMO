using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Threading;

namespace OfficeIMO.Html.Css;

/// <summary>Parses the selected provider-independent Selectors Level 4 subset.</summary>
public static class HtmlCssSelectorParser {
    /// <summary>
    /// Parses one complex selector. The current owned subset covers type, universal, id, class,
    /// attribute selectors and descendant, child, next-sibling and subsequent-sibling combinators.
    /// Selector lists, namespaces and pseudo selectors return <see cref="HtmlCssSelectorParseStatus.Unsupported"/>.
    /// </summary>
    public static HtmlCssSelectorParseResult Parse(
        string selector,
        HtmlCssSelectorOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (selector == null) throw new ArgumentNullException(nameof(selector));
        HtmlCssSelectorOptions effective = (options ?? new HtmlCssSelectorOptions()).Clone();
        effective.Validate();
        cancellationToken.ThrowIfCancellationRequested();
        string source = selector.Trim();
        IReadOnlyList<HtmlCssToken> tokens;
        try {
            tokens = HtmlCssTokenizer.Tokenize(source, new HtmlCssTokenizationOptions {
                MaxInputCharacters = effective.MaxInputCharacters,
                MaxTokens = effective.MaxTokens
            }, cancellationToken);
        } catch (HtmlCssTokenizationLimitException exception) {
            throw new HtmlCssSelectorLimitException(exception.LimitName, exception.Actual, exception.Maximum);
        }
        var parser = new Parser(source, tokens, effective, cancellationToken);
        return parser.Parse();
    }

    private sealed class Parser {
        private readonly string _source;
        private readonly IReadOnlyList<HtmlCssToken> _tokens;
        private readonly HtmlCssSelectorOptions _options;
        private readonly CancellationToken _cancellation;
        private int _position;
        private int _simpleSelectors;
        private int _ids;
        private int _classes;
        private int _types;

        internal Parser(string source, IReadOnlyList<HtmlCssToken> tokens, HtmlCssSelectorOptions options, CancellationToken cancellation) {
            _source = source; _tokens = tokens; _options = options; _cancellation = cancellation;
        }

        internal HtmlCssSelectorParseResult Parse() {
            SkipTrivia();
            if (Current.Kind == HtmlCssTokenKind.EndOfFile)
                return Invalid("empty-selector");
            var compounds = new List<HtmlCssSelectorCompound>();
            var combinators = new List<HtmlCssCombinator>();
            while (true) {
                _cancellation.ThrowIfCancellationRequested();
                ParseState state = ParseCompound(out HtmlCssSelectorCompound? compound, out string? reason);
                if (state != ParseState.Success)
                    return state == ParseState.Unsupported ? Unsupported(reason!) : Invalid(reason!);
                compounds.Add(compound!);
                if (compounds.Count > _options.MaxCompounds)
                    throw new HtmlCssSelectorLimitException(nameof(HtmlCssSelectorOptions.MaxCompounds), compounds.Count, _options.MaxCompounds);

                bool hadWhitespace = SkipTrivia();
                if (Current.Kind == HtmlCssTokenKind.EndOfFile) break;
                if (Current.Kind == HtmlCssTokenKind.Comma) return Unsupported("selector-list");
                HtmlCssCombinator combinator;
                if (IsDelimiter(">") || IsDelimiter("+") || IsDelimiter("~")) {
                    combinator = Current.Value == ">" ? HtmlCssCombinator.Child
                        : Current.Value == "+" ? HtmlCssCombinator.NextSibling
                        : HtmlCssCombinator.SubsequentSibling;
                    _position++;
                    SkipTrivia();
                } else if (hadWhitespace && StartsCompound(Current)) {
                    combinator = HtmlCssCombinator.Descendant;
                } else {
                    return Current.Kind == HtmlCssTokenKind.Colon || IsDelimiter("|") || IsDelimiter("&")
                        ? Unsupported("unsupported-selector-feature")
                        : Invalid("missing-combinator");
                }
                if (!StartsCompound(Current)) return Invalid("missing-compound-selector");
                combinators.Add(combinator);
            }
            var specificity = new HtmlCssSelectorSpecificity(_ids, _classes, _types);
            var parsed = new HtmlCssSelector(_source,
                new ReadOnlyCollection<HtmlCssSelectorCompound>(compounds),
                new ReadOnlyCollection<HtmlCssCombinator>(combinators), specificity);
            return new HtmlCssSelectorParseResult(HtmlCssSelectorParseStatus.Parsed, _source, parsed, null);
        }

        private ParseState ParseCompound(out HtmlCssSelectorCompound? compound, out string? reason) {
            compound = new HtmlCssSelectorCompound();
            reason = null;
            bool any = false;
            if (Current.Kind == HtmlCssTokenKind.Identifier) {
                compound.TypeName = Current.Value ?? string.Empty;
                _types++; _position++; any = true; RecordSimple();
            } else if (IsDelimiter("*")) {
                compound.Universal = true;
                _position++; any = true; RecordSimple();
            }
            while (true) {
                _cancellation.ThrowIfCancellationRequested();
                if (Current.Kind == HtmlCssTokenKind.Comment) { _position++; continue; }
                if (Current.Kind == HtmlCssTokenKind.Hash) {
                    if (!Current.IsIdentifierHash) { reason = "invalid-id-selector"; return ParseState.Invalid; }
                    compound.Ids.Add(Current.Value ?? string.Empty);
                    _ids++; _position++; any = true; RecordSimple();
                    continue;
                }
                if (IsDelimiter(".")) {
                    _position++;
                    if (Current.Kind != HtmlCssTokenKind.Identifier) { reason = "invalid-class-selector"; return ParseState.Invalid; }
                    compound.Classes.Add(Current.Value ?? string.Empty);
                    _classes++; _position++; any = true; RecordSimple();
                    continue;
                }
                if (Current.Kind == HtmlCssTokenKind.OpenBracket) {
                    ParseState attributeState = ParseAttribute(out HtmlCssAttributeSelector? attribute, out reason);
                    if (attributeState != ParseState.Success) return attributeState;
                    compound.Attributes.Add(attribute!);
                    _classes++; any = true; RecordSimple();
                    continue;
                }
                if (Current.Kind == HtmlCssTokenKind.Colon || IsDelimiter("|") || IsDelimiter("&")) {
                    reason = "unsupported-selector-feature";
                    return ParseState.Unsupported;
                }
                break;
            }
            if (!any) { reason = "expected-compound-selector"; return ParseState.Invalid; }
            return ParseState.Success;
        }

        private ParseState ParseAttribute(out HtmlCssAttributeSelector? selector, out string? reason) {
            selector = null; reason = null; _position++;
            SkipTrivia();
            if (Current.Kind != HtmlCssTokenKind.Identifier) { reason = "invalid-attribute-name"; return ParseState.Invalid; }
            string name = Current.Value ?? string.Empty;
            _position++; SkipTrivia();
            if (Current.Kind == HtmlCssTokenKind.CloseBracket) {
                _position++;
                selector = new HtmlCssAttributeSelector(name, HtmlCssAttributeOperator.Present, null, HtmlCssAttributeCase.Default);
                return ParseState.Success;
            }
            HtmlCssAttributeOperator operation;
            if (IsDelimiter("=")) operation = HtmlCssAttributeOperator.Equals;
            else if (IsAttributeOperator("~")) operation = HtmlCssAttributeOperator.Includes;
            else if (IsAttributeOperator("|")) operation = HtmlCssAttributeOperator.DashMatch;
            else if (IsAttributeOperator("^")) operation = HtmlCssAttributeOperator.Prefix;
            else if (IsAttributeOperator("$")) operation = HtmlCssAttributeOperator.Suffix;
            else if (IsAttributeOperator("*")) operation = HtmlCssAttributeOperator.Substring;
            else if (IsDelimiter("|")) { reason = "attribute-namespace"; return ParseState.Unsupported; }
            else { reason = "invalid-attribute-operator"; return ParseState.Invalid; }
            _position += operation == HtmlCssAttributeOperator.Equals ? 1 : 2;
            SkipTrivia();
            if (Current.Kind != HtmlCssTokenKind.Identifier && Current.Kind != HtmlCssTokenKind.String) {
                reason = "invalid-attribute-value"; return ParseState.Invalid;
            }
            if (Current.Kind == HtmlCssTokenKind.String && !IsClosedString(Current)) {
                reason = "unclosed-attribute-string"; return ParseState.Invalid;
            }
            string value = Current.Value ?? string.Empty;
            _position++;
            bool hadSpace = SkipTrivia();
            HtmlCssAttributeCase caseMode = HtmlCssAttributeCase.Default;
            if (hadSpace && Current.Kind == HtmlCssTokenKind.Identifier) {
                string modifier = Current.Value ?? string.Empty;
                if (string.Equals(modifier, "i", StringComparison.OrdinalIgnoreCase)) caseMode = HtmlCssAttributeCase.Insensitive;
                else if (string.Equals(modifier, "s", StringComparison.OrdinalIgnoreCase)) caseMode = HtmlCssAttributeCase.Sensitive;
                else {
                    reason = "unsupported-attribute-modifier"; return ParseState.Unsupported;
                }
                _position++; SkipTrivia();
            }
            if (Current.Kind != HtmlCssTokenKind.CloseBracket) { reason = "unclosed-attribute-selector"; return ParseState.Invalid; }
            _position++;
            selector = new HtmlCssAttributeSelector(name, operation, value, caseMode);
            return ParseState.Success;
        }

        private void RecordSimple() {
            _simpleSelectors++;
            if (_simpleSelectors > _options.MaxSimpleSelectors)
                throw new HtmlCssSelectorLimitException(nameof(HtmlCssSelectorOptions.MaxSimpleSelectors), _simpleSelectors, _options.MaxSimpleSelectors);
        }

        private bool SkipTrivia() {
            bool found = false;
            while (Current.Kind == HtmlCssTokenKind.Whitespace || Current.Kind == HtmlCssTokenKind.Comment) {
                if (Current.Kind == HtmlCssTokenKind.Whitespace) found = true;
                _position++; _cancellation.ThrowIfCancellationRequested();
            }
            return found;
        }

        private bool IsClosedString(HtmlCssToken token) {
            string text = token.GetText(_source);
            return text.Length >= 2 && (text[0] == '\'' || text[0] == '"') && text[text.Length - 1] == text[0];
        }

        private bool IsAttributeOperator(string prefix) => IsDelimiter(prefix) && Peek(1).Kind == HtmlCssTokenKind.Delimiter && Peek(1).Value == "=";
        private bool IsDelimiter(string value) => Current.Kind == HtmlCssTokenKind.Delimiter && Current.Value == value;
        private HtmlCssToken Current => Peek(0);
        private HtmlCssToken Peek(int offset) {
            int index = _position + offset;
            return index >= 0 && index < _tokens.Count ? _tokens[index] : _tokens[_tokens.Count - 1];
        }
        private static bool StartsCompound(HtmlCssToken token) => token.Kind == HtmlCssTokenKind.Identifier
            || token.Kind == HtmlCssTokenKind.Hash || token.Kind == HtmlCssTokenKind.OpenBracket
            || token.Kind == HtmlCssTokenKind.Delimiter && (token.Value == "*" || token.Value == ".");
        private HtmlCssSelectorParseResult Invalid(string reason) => new HtmlCssSelectorParseResult(HtmlCssSelectorParseStatus.InvalidSyntax, _source, null, reason);
        private HtmlCssSelectorParseResult Unsupported(string reason) => new HtmlCssSelectorParseResult(HtmlCssSelectorParseStatus.Unsupported, _source, null, reason);
        private enum ParseState { Success, Unsupported, Invalid }
    }
}
