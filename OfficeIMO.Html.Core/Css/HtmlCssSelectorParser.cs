using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Threading;

namespace OfficeIMO.Html.Css;

/// <summary>Parses the selected provider-independent Selectors Level 4 subset.</summary>
public static class HtmlCssSelectorParser {
    /// <summary>Parses one complex selector. A top-level comma is classified as a selector-list feature.</summary>
    public static HtmlCssSelectorParseResult Parse(string selector, HtmlCssSelectorOptions? options = null, CancellationToken cancellationToken = default) {
        Prepared prepared = Prepare(selector, options, cancellationToken);
        var parser = new Parser(prepared.Source, prepared.Tokens, prepared.Options, cancellationToken, retainProviderPseudoClasses: false);
        ParseOutcome outcome = parser.ParseSingle();
        return new HtmlCssSelectorParseResult(outcome.Status, prepared.Source, outcome.Selector, outcome.Reason);
    }

    internal static HtmlCssSelectorParseResult ParseHybrid(string selector, HtmlCssSelectorOptions? options = null, CancellationToken cancellationToken = default) {
        Prepared prepared = Prepare(selector, options, cancellationToken);
        var parser = new Parser(prepared.Source, prepared.Tokens, prepared.Options, cancellationToken, retainProviderPseudoClasses: true);
        ParseOutcome outcome = parser.ParseSingle();
        return new HtmlCssSelectorParseResult(outcome.Status, prepared.Source, outcome.Selector, outcome.Reason);
    }

    /// <summary>Parses a comma-separated selector list atomically. Every list member must belong to the owned subset.</summary>
    public static HtmlCssSelectorListParseResult ParseList(string selectorList, HtmlCssSelectorOptions? options = null, CancellationToken cancellationToken = default) {
        Prepared prepared = Prepare(selectorList, options, cancellationToken);
        var parser = new Parser(prepared.Source, prepared.Tokens, prepared.Options, cancellationToken, retainProviderPseudoClasses: false);
        ListOutcome outcome = parser.ParseList();
        HtmlCssSelectorList? list = outcome.Selectors == null ? null : new HtmlCssSelectorList(prepared.Source, new ReadOnlyCollection<HtmlCssSelector>(outcome.Selectors));
        return new HtmlCssSelectorListParseResult(outcome.Status, prepared.Source, list, outcome.Reason);
    }

    private static Prepared Prepare(string selector, HtmlCssSelectorOptions? options, CancellationToken cancellationToken) {
        if (selector == null) throw new ArgumentNullException(nameof(selector));
        HtmlCssSelectorOptions effective = (options ?? new HtmlCssSelectorOptions()).Clone();
        effective.Validate();
        cancellationToken.ThrowIfCancellationRequested();
        string source = selector.Trim();
        try {
            return new Prepared(source, HtmlCssTokenizer.Tokenize(source, new HtmlCssTokenizationOptions {
                MaxInputCharacters = effective.MaxInputCharacters,
                MaxTokens = effective.MaxTokens
            }, cancellationToken), effective);
        } catch (HtmlCssTokenizationLimitException exception) {
            throw new HtmlCssSelectorLimitException(exception.LimitName, exception.Actual, exception.Maximum);
        }
    }

    private readonly struct Prepared {
        internal Prepared(string source, IReadOnlyList<HtmlCssToken> tokens, HtmlCssSelectorOptions options) { Source = source; Tokens = tokens; Options = options; }
        internal string Source { get; }
        internal IReadOnlyList<HtmlCssToken> Tokens { get; }
        internal HtmlCssSelectorOptions Options { get; }
    }

    private sealed class Parser {
        private readonly string _source;
        private readonly IReadOnlyList<HtmlCssToken> _tokens;
        private readonly HtmlCssSelectorOptions _options;
        private readonly CancellationToken _cancellation;
        private readonly bool _retainProviderPseudoClasses;
        private int _position;
        private int _simpleSelectors;
        private int _selectorCount;
        private int _nestingDepth;

        internal Parser(string source, IReadOnlyList<HtmlCssToken> tokens, HtmlCssSelectorOptions options, CancellationToken cancellation,
            bool retainProviderPseudoClasses) {
            _source = source; _tokens = tokens; _options = options; _cancellation = cancellation;
            _retainProviderPseudoClasses = retainProviderPseudoClasses;
        }

        internal ParseOutcome ParseSingle() {
            SkipTrivia();
            if (Current.Kind == HtmlCssTokenKind.EndOfFile) return ParseOutcome.Invalid("empty-selector");
            ParseOutcome outcome = ParseComplexSelector();
            if (outcome.Status != HtmlCssSelectorParseStatus.Parsed) return outcome;
            SkipTrivia();
            if (Current.Kind == HtmlCssTokenKind.Comma) return ParseOutcome.Unsupported("selector-list");
            if (Current.Kind != HtmlCssTokenKind.EndOfFile) return ClassifyRemainder();
            return outcome;
        }

        internal ListOutcome ParseList() => ParseListUntil(
            HtmlCssTokenKind.EndOfFile,
            forgiving: false,
            suppressDefaultNamespaceOnSubject: false);

        private ListOutcome ParseListUntil(
            HtmlCssTokenKind terminator,
            bool forgiving,
            bool suppressDefaultNamespaceOnSubject) {
            SkipTrivia();
            if (Current.Kind == terminator) return ListOutcome.Invalid("empty-selector-list");
            var selectors = new List<HtmlCssSelector>();
            string? unsupportedReason = null;
            string? invalidReason = null;
            while (true) {
                _cancellation.ThrowIfCancellationRequested();
                ParseOutcome outcome = ParseComplexSelector(suppressDefaultNamespaceOnSubject);
                if (outcome.Status == HtmlCssSelectorParseStatus.Parsed) selectors.Add(outcome.Selector!);
                else {
                    bool balanced = SkipToListBoundary(terminator);
                    if (!balanced) invalidReason ??= "unclosed-selector-list-item";
                    else if (outcome.Status == HtmlCssSelectorParseStatus.Unsupported) unsupportedReason ??= outcome.Reason;
                    else if (!forgiving) invalidReason ??= outcome.Reason;
                }
                RecordSelector();
                SkipTrivia();
                if (Current.Kind == terminator) break;
                if (Current.Kind != HtmlCssTokenKind.Comma) {
                    invalidReason ??= "missing-list-separator";
                    if (!SkipToListBoundary(terminator)) invalidReason ??= "unclosed-selector-list-item";
                    if (Current.Kind == terminator) break;
                }
                _position++;
                SkipTrivia();
                if (Current.Kind == terminator || Current.Kind == HtmlCssTokenKind.EndOfFile) {
                    if (!forgiving) invalidReason ??= "empty-selector-list-item";
                    break;
                }
            }
            if (invalidReason != null) return ListOutcome.Invalid(invalidReason);
            if (unsupportedReason != null) return new ListOutcome(HtmlCssSelectorParseStatus.Unsupported, null, unsupportedReason);
            if (selectors.Count == 0) return ListOutcome.Invalid("empty-selector-list");
            return new ListOutcome(HtmlCssSelectorParseStatus.Parsed, selectors, null);
        }

        private bool SkipToListBoundary(HtmlCssTokenKind terminator) {
            int parentheses = 0;
            int brackets = 0;
            while (Current.Kind != HtmlCssTokenKind.EndOfFile) {
                _cancellation.ThrowIfCancellationRequested();
                HtmlCssTokenKind kind = Current.Kind;
                if (kind == terminator && parentheses == 0 && brackets == 0) return true;
                if (kind == HtmlCssTokenKind.Comma && parentheses == 0 && brackets == 0) return true;
                if (kind == HtmlCssTokenKind.Function || kind == HtmlCssTokenKind.OpenParenthesis) parentheses++;
                else if (kind == HtmlCssTokenKind.OpenBracket) brackets++;
                else if (kind == HtmlCssTokenKind.CloseParenthesis) {
                    if (parentheses == 0) return terminator == HtmlCssTokenKind.CloseParenthesis;
                    parentheses--;
                } else if (kind == HtmlCssTokenKind.CloseBracket) {
                    if (brackets == 0) return false;
                    brackets--;
                }
                _position++;
            }
            return terminator == HtmlCssTokenKind.EndOfFile && parentheses == 0 && brackets == 0;
        }

        private ParseOutcome ParseComplexSelector(bool suppressDefaultNamespaceOnSubject = false) {
            int startOffset = Current.Offset;
            var compounds = new List<HtmlCssSelectorCompound>();
            var combinators = new List<HtmlCssCombinator>();
            var specificity = new HtmlCssSelectorSpecificity(0, 0, 0);
            while (true) {
                _cancellation.ThrowIfCancellationRequested();
                ParseState state = ParseCompound(out HtmlCssSelectorCompound? compound, out HtmlCssSelectorSpecificity compoundSpecificity, out string? reason);
                if (state != ParseState.Success) return state == ParseState.Unsupported ? ParseOutcome.Unsupported(reason!) : ParseOutcome.Invalid(reason!);
                compounds.Add(compound!);
                specificity = HtmlCssSelectorSpecificity.Add(specificity, compoundSpecificity);
                if (compounds.Count > _options.MaxCompounds) throw new HtmlCssSelectorLimitException(nameof(HtmlCssSelectorOptions.MaxCompounds), compounds.Count, _options.MaxCompounds);
                bool hadWhitespace = SkipTrivia();
                if (Current.Kind == HtmlCssTokenKind.EndOfFile || Current.Kind == HtmlCssTokenKind.Comma || Current.Kind == HtmlCssTokenKind.CloseParenthesis) break;
                HtmlCssCombinator combinator;
                if (IsDelimiter(">") || IsDelimiter("+") || IsDelimiter("~")) {
                    combinator = Current.Value == ">" ? HtmlCssCombinator.Child : Current.Value == "+" ? HtmlCssCombinator.NextSibling : HtmlCssCombinator.SubsequentSibling;
                    _position++;
                    SkipTrivia();
                } else if (hadWhitespace && StartsCompound(Current)) combinator = HtmlCssCombinator.Descendant;
                else return ClassifyRemainder();
                if (!StartsCompound(Current)) return ParseOutcome.Invalid("missing-compound-selector");
                combinators.Add(combinator);
            }
            if (suppressDefaultNamespaceOnSubject && compounds.Count > 0 && compounds[compounds.Count - 1].HasImplicitDefaultNamespace) {
                compounds[compounds.Count - 1].Namespace = HtmlCssNamespaceConstraint.Any;
                compounds[compounds.Count - 1].HasImplicitDefaultNamespace = false;
            }
            int endOffset = PreviousSignificantEnd(startOffset);
            string source = endOffset <= startOffset ? string.Empty : _source.Substring(startOffset, endOffset - startOffset).Trim();
            return ParseOutcome.Parsed(new HtmlCssSelector(source, new ReadOnlyCollection<HtmlCssSelectorCompound>(compounds), new ReadOnlyCollection<HtmlCssCombinator>(combinators), specificity));
        }

        private ParseState ParseCompound(out HtmlCssSelectorCompound? compound, out HtmlCssSelectorSpecificity specificity, out string? reason) {
            compound = new HtmlCssSelectorCompound();
            specificity = new HtmlCssSelectorSpecificity(0, 0, 0);
            reason = null;
            bool any = false;
            ParseState typeState = ParseTypeOrUniversal(compound, ref specificity, ref any, out reason);
            if (typeState != ParseState.Success) return typeState;
            if (!any && _options.Namespaces?.DefaultNamespaceUri != null) {
                compound.Namespace = HtmlCssNamespaceConstraint.Exact(_options.Namespaces.DefaultNamespaceUri);
                compound.HasImplicitDefaultNamespace = true;
            }
            while (true) {
                _cancellation.ThrowIfCancellationRequested();
                if (Current.Kind == HtmlCssTokenKind.Comment) { _position++; continue; }
                if (Current.Kind == HtmlCssTokenKind.Hash) {
                    if (!Current.IsIdentifierHash) { reason = "invalid-id-selector"; return ParseState.Invalid; }
                    compound.Ids.Add(Current.Value ?? string.Empty);
                    specificity = HtmlCssSelectorSpecificity.Add(specificity, new HtmlCssSelectorSpecificity(1, 0, 0));
                    _position++; any = true; RecordSimple(); continue;
                }
                if (IsDelimiter(".")) {
                    _position++;
                    if (Current.Kind != HtmlCssTokenKind.Identifier) { reason = "invalid-class-selector"; return ParseState.Invalid; }
                    compound.Classes.Add(Current.Value ?? string.Empty);
                    specificity = HtmlCssSelectorSpecificity.Add(specificity, new HtmlCssSelectorSpecificity(0, 1, 0));
                    _position++; any = true; RecordSimple(); continue;
                }
                if (Current.Kind == HtmlCssTokenKind.OpenBracket) {
                    ParseState attributeState = ParseAttribute(out HtmlCssAttributeSelector? attribute, out reason);
                    if (attributeState != ParseState.Success) return attributeState;
                    compound.Attributes.Add(attribute!);
                    specificity = HtmlCssSelectorSpecificity.Add(specificity, new HtmlCssSelectorSpecificity(0, 1, 0));
                    any = true; RecordSimple(); continue;
                }
                if (Current.Kind == HtmlCssTokenKind.Colon) {
                    ParseState pseudoState = ParsePseudoClass(out HtmlCssPseudoClassSelector? pseudo, out HtmlCssSelectorSpecificity pseudoSpecificity, out reason);
                    if (pseudoState != ParseState.Success) return pseudoState;
                    compound.PseudoClasses.Add(pseudo!);
                    specificity = HtmlCssSelectorSpecificity.Add(specificity, pseudoSpecificity);
                    any = true; RecordSimple(); continue;
                }
                if (IsDelimiter("&")) { reason = "nesting-selector"; return ParseState.Unsupported; }
                break;
            }
            if (!any) { reason = "expected-compound-selector"; return ParseState.Invalid; }
            return ParseState.Success;
        }

        private ParseState ParseTypeOrUniversal(HtmlCssSelectorCompound compound, ref HtmlCssSelectorSpecificity specificity, ref bool any, out string? reason) {
            reason = null;
            if (Current.Kind != HtmlCssTokenKind.Identifier && !IsDelimiter("*") && !IsDelimiter("|")) return ParseState.Success;
            HtmlCssNamespaceConstraint ns;
            string? localName = null;
            bool universal = false;
            if (IsDelimiter("|")) {
                ns = HtmlCssNamespaceConstraint.None; _position++; SkipComments();
                if (!ReadLocalName(out localName, out universal)) { reason = "invalid-qualified-type-selector"; return ParseState.Invalid; }
            } else {
                HtmlCssToken first = Current;
                int pipeIndex = NextNonCommentIndex(_position + 1);
                bool prefixCandidate = TokenAt(pipeIndex).Kind == HtmlCssTokenKind.Delimiter && TokenAt(pipeIndex).Value == "|";
                if (prefixCandidate) {
                    _position = pipeIndex + 1;
                    SkipComments();
                    if (first.Kind == HtmlCssTokenKind.Identifier) {
                        if (_options.Namespaces == null || !_options.Namespaces.TryGetNamespaceUri(first.Value ?? string.Empty, out string uri)) { reason = "undeclared-namespace-prefix"; return ParseState.Invalid; }
                        ns = HtmlCssNamespaceConstraint.Exact(uri);
                    } else ns = HtmlCssNamespaceConstraint.Any;
                    if (!ReadLocalName(out localName, out universal)) { reason = "invalid-qualified-type-selector"; return ParseState.Invalid; }
                } else {
                    ns = _options.Namespaces?.DefaultNamespaceUri == null ? HtmlCssNamespaceConstraint.Any : HtmlCssNamespaceConstraint.Exact(_options.Namespaces.DefaultNamespaceUri);
                    if (first.Kind == HtmlCssTokenKind.Identifier) localName = first.Value ?? string.Empty; else universal = true;
                    _position++;
                }
            }
            compound.Namespace = ns; compound.TypeName = localName; compound.Universal = universal;
            specificity = HtmlCssSelectorSpecificity.Add(specificity, new HtmlCssSelectorSpecificity(0, 0, universal ? 0 : 1));
            any = true; RecordSimple();
            return ParseState.Success;
        }

        private bool ReadLocalName(out string? localName, out bool universal) {
            localName = null; universal = false;
            if (Current.Kind == HtmlCssTokenKind.Identifier) { localName = Current.Value ?? string.Empty; _position++; return true; }
            if (IsDelimiter("*")) { universal = true; _position++; return true; }
            return false;
        }

        private ParseState ParseAttribute(out HtmlCssAttributeSelector? selector, out string? reason) {
            selector = null; reason = null; _position++; SkipTrivia();
            HtmlCssNamespaceConstraint ns = HtmlCssNamespaceConstraint.None;
            string? name;
            if (IsDelimiter("|")) {
                _position++; SkipComments();
                if (Current.Kind != HtmlCssTokenKind.Identifier) { reason = "invalid-attribute-name"; return ParseState.Invalid; }
                name = Current.Value ?? string.Empty; _position++;
            } else if (Current.Kind == HtmlCssTokenKind.Identifier || IsDelimiter("*")) {
                int pipeIndex = NextNonCommentIndex(_position + 1);
                int afterPipeIndex = NextNonCommentIndex(pipeIndex + 1);
                bool qualified = TokenAt(pipeIndex).Kind == HtmlCssTokenKind.Delimiter && TokenAt(pipeIndex).Value == "|"
                    && !(TokenAt(afterPipeIndex).Kind == HtmlCssTokenKind.Delimiter && TokenAt(afterPipeIndex).Value == "=");
                if (!qualified) {
                    if (Current.Kind != HtmlCssTokenKind.Identifier) { reason = "invalid-attribute-name"; return ParseState.Invalid; }
                    name = Current.Value ?? string.Empty; _position++;
                } else {
                HtmlCssToken prefix = Current; _position = pipeIndex + 1; SkipComments();
                if (prefix.Kind == HtmlCssTokenKind.Identifier) {
                    if (_options.Namespaces == null || !_options.Namespaces.TryGetNamespaceUri(prefix.Value ?? string.Empty, out string uri)) { reason = "undeclared-namespace-prefix"; return ParseState.Invalid; }
                    ns = HtmlCssNamespaceConstraint.Exact(uri);
                } else ns = HtmlCssNamespaceConstraint.Any;
                if (Current.Kind != HtmlCssTokenKind.Identifier) { reason = "invalid-attribute-name"; return ParseState.Invalid; }
                name = Current.Value ?? string.Empty; _position++;
                }
            }
            else { reason = "invalid-attribute-name"; return ParseState.Invalid; }
            SkipTrivia();
            if (Current.Kind == HtmlCssTokenKind.CloseBracket) {
                _position++;
                selector = new HtmlCssAttributeSelector(name, ns, HtmlCssAttributeOperator.Present, null, HtmlCssAttributeCase.Default);
                return ParseState.Success;
            }
            HtmlCssAttributeOperator operation;
            if (IsDelimiter("=")) operation = HtmlCssAttributeOperator.Equals;
            else if (IsAttributeOperator("~")) operation = HtmlCssAttributeOperator.Includes;
            else if (IsAttributeOperator("|")) operation = HtmlCssAttributeOperator.DashMatch;
            else if (IsAttributeOperator("^")) operation = HtmlCssAttributeOperator.Prefix;
            else if (IsAttributeOperator("$")) operation = HtmlCssAttributeOperator.Suffix;
            else if (IsAttributeOperator("*")) operation = HtmlCssAttributeOperator.Substring;
            else { reason = "invalid-attribute-operator"; return ParseState.Invalid; }
            _position += operation == HtmlCssAttributeOperator.Equals ? 1 : 2; SkipTrivia();
            if (Current.Kind != HtmlCssTokenKind.Identifier && Current.Kind != HtmlCssTokenKind.String) { reason = "invalid-attribute-value"; return ParseState.Invalid; }
            if (Current.Kind == HtmlCssTokenKind.String && !IsClosedString(Current)) { reason = "unclosed-attribute-string"; return ParseState.Invalid; }
            string value = Current.Value ?? string.Empty; _position++;
            bool hadSpace = SkipTrivia();
            HtmlCssAttributeCase caseMode = HtmlCssAttributeCase.Default;
            if (hadSpace && Current.Kind == HtmlCssTokenKind.Identifier) {
                string modifier = Current.Value ?? string.Empty;
                if (string.Equals(modifier, "i", StringComparison.OrdinalIgnoreCase)) caseMode = HtmlCssAttributeCase.Insensitive;
                else if (string.Equals(modifier, "s", StringComparison.OrdinalIgnoreCase)) caseMode = HtmlCssAttributeCase.Sensitive;
                else { reason = "unsupported-attribute-modifier"; return ParseState.Unsupported; }
                _position++; SkipTrivia();
            }
            if (Current.Kind != HtmlCssTokenKind.CloseBracket) { reason = "unclosed-attribute-selector"; return ParseState.Invalid; }
            _position++;
            selector = new HtmlCssAttributeSelector(name, ns, operation, value, caseMode);
            return ParseState.Success;
        }

        private ParseState ParsePseudoClass(out HtmlCssPseudoClassSelector? selector, out HtmlCssSelectorSpecificity specificity, out string? reason) {
            int pseudoStart = Current.Offset;
            selector = null; reason = null; specificity = new HtmlCssSelectorSpecificity(0, 1, 0); _position++;
            if (Current.Kind == HtmlCssTokenKind.Colon) { reason = "pseudo-element"; return ParseState.Unsupported; }
            if (Current.Kind == HtmlCssTokenKind.Identifier) {
                string name = (Current.Value ?? string.Empty).ToLowerInvariant(); _position++;
                HtmlCssPseudoClassKind? kind = name switch {
                    "root" => HtmlCssPseudoClassKind.Root, "empty" => HtmlCssPseudoClassKind.Empty,
                    "first-child" => HtmlCssPseudoClassKind.FirstChild, "last-child" => HtmlCssPseudoClassKind.LastChild,
                    "only-child" => HtmlCssPseudoClassKind.OnlyChild, "first-of-type" => HtmlCssPseudoClassKind.FirstOfType,
                    "last-of-type" => HtmlCssPseudoClassKind.LastOfType, "only-of-type" => HtmlCssPseudoClassKind.OnlyOfType, _ => null
                };
                if (!kind.HasValue) {
                    if (_retainProviderPseudoClasses) {
                        int end = _tokens[_position - 1].Offset + _tokens[_position - 1].Length;
                        selector = new HtmlCssPseudoClassSelector(HtmlCssPseudoClassKind.Provider,
                            providerSource: _source.Substring(pseudoStart, end - pseudoStart));
                        return ParseState.Success;
                    }
                    reason = "unsupported-pseudo-class"; return ParseState.Unsupported;
                }
                selector = new HtmlCssPseudoClassSelector(kind.Value); return ParseState.Success;
            }
            if (Current.Kind != HtmlCssTokenKind.Function) { reason = "invalid-pseudo-class"; return ParseState.Invalid; }
            string function = (Current.Value ?? string.Empty).ToLowerInvariant();
            if (function == "lang") {
                _position++;
                SkipTrivia();
                if (Current.Kind != HtmlCssTokenKind.Identifier) {
                    reason = "unsupported-lang-range"; return ParseState.Unsupported;
                }
                string range = Current.Value ?? string.Empty;
                if (!IsSupportedBasicLanguageRange(range)) {
                    reason = "unsupported-lang-range"; return ParseState.Unsupported;
                }
                _position++;
                SkipTrivia();
                if (Current.Kind != HtmlCssTokenKind.CloseParenthesis) {
                    reason = Current.Kind == HtmlCssTokenKind.Comma ? "lang-range-list" : "invalid-lang-range";
                    return Current.Kind == HtmlCssTokenKind.Comma ? ParseState.Unsupported : ParseState.Invalid;
                }
                _position++;
                selector = new HtmlCssPseudoClassSelector(HtmlCssPseudoClassKind.Lang, stringArgument: range);
                return ParseState.Success;
            }
            if (function == "is" || function == "where" || function == "not") {
                if (++_nestingDepth > _options.MaxNestingDepth) throw new HtmlCssSelectorLimitException(nameof(HtmlCssSelectorOptions.MaxNestingDepth), _nestingDepth, _options.MaxNestingDepth);
                _position++;
                ListOutcome nested = ParseListUntil(
                    HtmlCssTokenKind.CloseParenthesis,
                    forgiving: function != "not",
                    suppressDefaultNamespaceOnSubject: true);
                _nestingDepth--;
                if (nested.Status != HtmlCssSelectorParseStatus.Parsed) { reason = nested.Reason; return nested.Status == HtmlCssSelectorParseStatus.Unsupported ? ParseState.Unsupported : ParseState.Invalid; }
                if (Current.Kind != HtmlCssTokenKind.CloseParenthesis) { reason = "unclosed-pseudo-class"; return ParseState.Invalid; }
                _position++;
                HtmlCssPseudoClassKind kind = function == "is" ? HtmlCssPseudoClassKind.Is : function == "where" ? HtmlCssPseudoClassKind.Where : HtmlCssPseudoClassKind.Not;
                specificity = function == "where" ? new HtmlCssSelectorSpecificity(0, 0, 0) : MaxSpecificity(nested.Selectors!);
                selector = new HtmlCssPseudoClassSelector(kind, selectors: new ReadOnlyCollection<HtmlCssSelector>(nested.Selectors!));
                return ParseState.Success;
            }
            HtmlCssPseudoClassKind? nthKind = function switch {
                "nth-child" => HtmlCssPseudoClassKind.NthChild, "nth-last-child" => HtmlCssPseudoClassKind.NthLastChild,
                "nth-of-type" => HtmlCssPseudoClassKind.NthOfType, "nth-last-of-type" => HtmlCssPseudoClassKind.NthLastOfType, _ => null
            };
            if (!nthKind.HasValue) {
                if (_retainProviderPseudoClasses) {
                    _position++;
                    int providerClose = FindFunctionClose();
                    if (providerClose < 0) { reason = "unclosed-pseudo-class"; return ParseState.Invalid; }
                    int end = _tokens[providerClose].Offset + _tokens[providerClose].Length;
                    selector = new HtmlCssPseudoClassSelector(HtmlCssPseudoClassKind.Provider,
                        providerSource: _source.Substring(pseudoStart, end - pseudoStart));
                    _position = providerClose + 1;
                    return ParseState.Success;
                }
                reason = "unsupported-pseudo-class"; return ParseState.Unsupported;
            }
            _position++;
            int bodyStart = Current.Offset;
            int close = FindFunctionClose();
            if (close < 0) { reason = "unclosed-pseudo-class"; return ParseState.Invalid; }
            for (int index = _position; index < close; index++) {
                if (_tokens[index].Kind == HtmlCssTokenKind.Identifier && string.Equals(_tokens[index].Value, "of", StringComparison.OrdinalIgnoreCase)) {
                    _position = close + 1; reason = "nth-selector-list"; return ParseState.Unsupported;
                }
                if (_tokens[index].Kind == HtmlCssTokenKind.Function || _tokens[index].Kind == HtmlCssTokenKind.OpenParenthesis) { reason = "invalid-nth-expression"; return ParseState.Invalid; }
            }
            int bodyEnd = _tokens[close].Offset;
            string expression = bodyEnd <= bodyStart ? string.Empty : _source.Substring(bodyStart, bodyEnd - bodyStart);
            if (!TryNormalizeAnPlusB(expression, out string normalized)
                || !TryParseAnPlusB(normalized, out HtmlCssAnPlusB formula)) { reason = "invalid-nth-expression"; return ParseState.Invalid; }
            _position = close + 1;
            selector = new HtmlCssPseudoClassSelector(nthKind.Value, formula); return ParseState.Success;
        }

        private int FindFunctionClose() {
            int depth = 1;
            for (int index = _position; index < _tokens.Count; index++) {
                HtmlCssTokenKind kind = _tokens[index].Kind;
                if (kind == HtmlCssTokenKind.Function || kind == HtmlCssTokenKind.OpenParenthesis) depth++;
                else if (kind == HtmlCssTokenKind.CloseParenthesis && --depth == 0) return index;
                else if (kind == HtmlCssTokenKind.EndOfFile) return -1;
            }
            return -1;
        }

        private static bool IsSupportedBasicLanguageRange(string range) {
            int subtagLength = 0;
            for (int index = 0; index < range.Length; index++) {
                char value = range[index];
                if (value == '-') {
                    if (subtagLength == 0 || subtagLength > 8) return false;
                    subtagLength = 0;
                } else if ((value >= 'A' && value <= 'Z') || (value >= 'a' && value <= 'z')
                    || (index > 0 && value >= '0' && value <= '9')) {
                    subtagLength++;
                } else return false;
            }
            return subtagLength > 0 && subtagLength <= 8;
        }

        private static bool TryParseAnPlusB(string expression, out HtmlCssAnPlusB formula) {
            formula = default;
            string value = expression.ToLowerInvariant();
            if (value == "odd") { formula = new HtmlCssAnPlusB(2, 1); return true; }
            if (value == "even") { formula = new HtmlCssAnPlusB(2, 0); return true; }
            int n = value.IndexOf('n');
            if (n < 0) {
                if (!int.TryParse(value, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out int b)) return false;
                formula = new HtmlCssAnPlusB(0, b); return true;
            }
            if (value.IndexOf('n', n + 1) >= 0) return false;
            string aText = value.Substring(0, n);
            int a;
            if (aText.Length == 0 || aText == "+") a = 1;
            else if (aText == "-") a = -1;
            else if (!int.TryParse(aText, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out a)) return false;
            string bText = value.Substring(n + 1);
            int bValue = 0;
            if (bText.Length != 0 && !int.TryParse(bText, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out bValue)) return false;
            formula = new HtmlCssAnPlusB(a, bValue); return true;
        }

        private static bool TryNormalizeAnPlusB(string value, out string normalized) {
            var retained = new System.Text.StringBuilder(value.Length);
            bool comment = false;
            for (int index = 0; index < value.Length; index++) {
                if (!comment && value[index] == '/' && index + 1 < value.Length && value[index + 1] == '*') { comment = true; retained.Append(' '); index++; continue; }
                if (comment && value[index] == '*' && index + 1 < value.Length && value[index + 1] == '/') { comment = false; index++; continue; }
                if (!comment) retained.Append(value[index]);
            }
            if (comment) { normalized = string.Empty; return false; }
            string spaced = retained.ToString().Trim();
            if (System.Text.RegularExpressions.Regex.IsMatch(spaced, @"^[+-]\s")
                || System.Text.RegularExpressions.Regex.IsMatch(spaced, @"\d\s+n", System.Text.RegularExpressions.RegexOptions.IgnoreCase)) {
                normalized = string.Empty; return false;
            }
            int n = spaced.IndexOf("n", StringComparison.OrdinalIgnoreCase);
            if (n >= 0) {
                string suffix = spaced.Substring(n + 1).TrimStart();
                if (suffix.Length > 0 && suffix[0] != '+' && suffix[0] != '-') {
                    normalized = string.Empty; return false;
                }
            }
            normalized = System.Text.RegularExpressions.Regex.Replace(spaced, @"\s+", string.Empty);
            return normalized.Length != 0;
        }

        private static HtmlCssSelectorSpecificity MaxSpecificity(IReadOnlyList<HtmlCssSelector> selectors) {
            var max = new HtmlCssSelectorSpecificity(0, 0, 0);
            foreach (HtmlCssSelector selector in selectors) max = HtmlCssSelectorSpecificity.Max(max, selector.Specificity);
            return max;
        }

        private void RecordSimple() {
            if (++_simpleSelectors > _options.MaxSimpleSelectors) throw new HtmlCssSelectorLimitException(nameof(HtmlCssSelectorOptions.MaxSimpleSelectors), _simpleSelectors, _options.MaxSimpleSelectors);
        }
        private void RecordSelector() {
            if (++_selectorCount > _options.MaxSelectors) throw new HtmlCssSelectorLimitException(nameof(HtmlCssSelectorOptions.MaxSelectors), _selectorCount, _options.MaxSelectors);
        }
        private bool SkipTrivia() {
            bool foundWhitespace = false;
            while (Current.Kind == HtmlCssTokenKind.Whitespace || Current.Kind == HtmlCssTokenKind.Comment) {
                if (Current.Kind == HtmlCssTokenKind.Whitespace) foundWhitespace = true;
                _position++; _cancellation.ThrowIfCancellationRequested();
            }
            return foundWhitespace;
        }
        private void SkipComments() {
            while (Current.Kind == HtmlCssTokenKind.Comment) {
                _position++;
                _cancellation.ThrowIfCancellationRequested();
            }
        }
        private int NextNonCommentIndex(int index) {
            while (TokenAt(index).Kind == HtmlCssTokenKind.Comment) index++;
            return index;
        }
        private int PreviousSignificantEnd(int fallback) {
            int index = _position - 1;
            while (index >= 0 && (_tokens[index].Kind == HtmlCssTokenKind.Whitespace || _tokens[index].Kind == HtmlCssTokenKind.Comment)) index--;
            return index < 0 ? fallback : _tokens[index].Offset + _tokens[index].Length;
        }
        private bool IsClosedString(HtmlCssToken token) {
            string text = token.GetText(_source);
            return text.Length >= 2 && (text[0] == '\'' || text[0] == '"') && text[text.Length - 1] == text[0];
        }
        private ParseOutcome ClassifyRemainder() => Current.Kind == HtmlCssTokenKind.Colon || Current.Kind == HtmlCssTokenKind.Function || IsDelimiter("|") || IsDelimiter("&")
            ? ParseOutcome.Unsupported("unsupported-selector-feature") : ParseOutcome.Invalid("missing-combinator");
        private bool IsAttributeOperator(string prefix) => IsDelimiter(prefix) && Peek(1).Kind == HtmlCssTokenKind.Delimiter && Peek(1).Value == "=";
        private bool IsDelimiter(string value) => Current.Kind == HtmlCssTokenKind.Delimiter && Current.Value == value;
        private HtmlCssToken Current => Peek(0);
        private HtmlCssToken Peek(int offset) {
            int index = _position + offset;
            return TokenAt(index);
        }
        private HtmlCssToken TokenAt(int index) => index >= 0 && index < _tokens.Count ? _tokens[index] : _tokens[_tokens.Count - 1];
        private static bool StartsCompound(HtmlCssToken token) => token.Kind == HtmlCssTokenKind.Identifier || token.Kind == HtmlCssTokenKind.Hash
            || token.Kind == HtmlCssTokenKind.OpenBracket || token.Kind == HtmlCssTokenKind.Colon
            || token.Kind == HtmlCssTokenKind.Delimiter && (token.Value == "*" || token.Value == "." || token.Value == "|");
    }

    private readonly struct ParseOutcome {
        private ParseOutcome(HtmlCssSelectorParseStatus status, HtmlCssSelector? selector, string? reason) { Status = status; Selector = selector; Reason = reason; }
        internal HtmlCssSelectorParseStatus Status { get; }
        internal HtmlCssSelector? Selector { get; }
        internal string? Reason { get; }
        internal static ParseOutcome Parsed(HtmlCssSelector selector) => new ParseOutcome(HtmlCssSelectorParseStatus.Parsed, selector, null);
        internal static ParseOutcome Unsupported(string reason) => new ParseOutcome(HtmlCssSelectorParseStatus.Unsupported, null, reason);
        internal static ParseOutcome Invalid(string reason) => new ParseOutcome(HtmlCssSelectorParseStatus.InvalidSyntax, null, reason);
    }

    private readonly struct ListOutcome {
        internal ListOutcome(HtmlCssSelectorParseStatus status, List<HtmlCssSelector>? selectors, string? reason) { Status = status; Selectors = selectors; Reason = reason; }
        internal HtmlCssSelectorParseStatus Status { get; }
        internal List<HtmlCssSelector>? Selectors { get; }
        internal string? Reason { get; }
        internal static ListOutcome Invalid(string reason) => new ListOutcome(HtmlCssSelectorParseStatus.InvalidSyntax, null, reason);
    }

    private enum ParseState { Success, Unsupported, Invalid }
}
