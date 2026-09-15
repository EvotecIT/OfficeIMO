using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Html.Css;

/// <summary>Parses provider-independent CSS rules, declarations, and nested component values while retaining exact source.</summary>
public static class HtmlCssSyntaxParser {
    /// <summary>
    /// Parses a stylesheet without validating selector or property grammars. Unknown at-rules,
    /// declarations, comments, whitespace, and recovered invalid source remain available.
    /// </summary>
    public static HtmlCssStyleSheet ParseStyleSheet(string source, HtmlCssSyntaxOptions? options = null, CancellationToken cancellationToken = default) {
        ParseResult parsed = Parse(source, options, cancellationToken, topLevel: true);
        return new HtmlCssStyleSheet(source, parsed.Tokens, parsed.Contents, parsed.Diagnostics);
    }

    /// <summary>
    /// Parses the contents of a style block, such as a <c>style</c> attribute, preserving duplicate
    /// and unknown declarations together with nested rules and invalid recovered source.
    /// </summary>
    public static HtmlCssStyleBlock ParseStyleBlock(string source, HtmlCssSyntaxOptions? options = null, CancellationToken cancellationToken = default) {
        ParseResult parsed = Parse(source, options, cancellationToken, topLevel: false);
        return new HtmlCssStyleBlock(source, parsed.Tokens, parsed.Contents, parsed.Diagnostics);
    }

    private static ParseResult Parse(string source, HtmlCssSyntaxOptions? options, CancellationToken cancellationToken, bool topLevel) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        HtmlCssSyntaxOptions effective = (options ?? new HtmlCssSyntaxOptions()).Clone();
        effective.Validate();
        cancellationToken.ThrowIfCancellationRequested();
        IReadOnlyList<HtmlCssToken> tokens;
        try {
            tokens = HtmlCssTokenizer.Tokenize(source, new HtmlCssTokenizationOptions {
                MaxInputCharacters = effective.MaxInputCharacters,
                MaxTokens = effective.MaxTokens
            }, cancellationToken);
        } catch (HtmlCssTokenizationLimitException exception) {
            throw new HtmlCssSyntaxLimitException(exception.LimitName, exception.Actual, exception.Maximum);
        }
        var context = new ParseContext(source, tokens, effective, cancellationToken);
        IReadOnlyList<HtmlCssSyntaxNode> contents = new Parser(context, 0, tokens.Count - 1, 0).ParseContents(topLevel);
        cancellationToken.ThrowIfCancellationRequested();
        return new ParseResult(tokens, contents, context.Diagnostics);
    }

    private readonly struct ParseResult {
        internal ParseResult(IReadOnlyList<HtmlCssToken> tokens, IReadOnlyList<HtmlCssSyntaxNode> contents, IReadOnlyList<HtmlCssSyntaxDiagnostic> diagnostics) {
            Tokens = tokens; Contents = contents; Diagnostics = diagnostics;
        }
        internal IReadOnlyList<HtmlCssToken> Tokens { get; }
        internal IReadOnlyList<HtmlCssSyntaxNode> Contents { get; }
        internal IReadOnlyList<HtmlCssSyntaxDiagnostic> Diagnostics { get; }
    }

    private sealed class ParseContext {
        internal ParseContext(string source, IReadOnlyList<HtmlCssToken> tokens, HtmlCssSyntaxOptions options, CancellationToken cancellation) {
            Source = source; Tokens = tokens; Options = options; Cancellation = cancellation;
        }
        internal string Source { get; }
        internal IReadOnlyList<HtmlCssToken> Tokens { get; }
        internal HtmlCssSyntaxOptions Options { get; }
        internal CancellationToken Cancellation { get; }
        internal List<HtmlCssSyntaxDiagnostic> MutableDiagnostics { get; } = new List<HtmlCssSyntaxDiagnostic>();
        internal IReadOnlyList<HtmlCssSyntaxDiagnostic> Diagnostics => MutableDiagnostics.AsReadOnly();
        private int _syntaxNodes;
        internal void RecordNode() {
            Cancellation.ThrowIfCancellationRequested();
            int actual = checked(++_syntaxNodes);
            if (Options.MaxSyntaxNodes.HasValue && actual > Options.MaxSyntaxNodes.Value)
                throw new HtmlCssSyntaxLimitException(nameof(Options.MaxSyntaxNodes), actual, Options.MaxSyntaxNodes.Value);
        }
        internal void CheckDepth(int depth, HtmlCssToken token) {
            if (Options.MaxNestingDepth.HasValue && depth > Options.MaxNestingDepth.Value)
                throw new HtmlCssSyntaxLimitException(nameof(Options.MaxNestingDepth), depth, Options.MaxNestingDepth.Value);
            Cancellation.ThrowIfCancellationRequested();
        }
        internal void Diagnose(string code, string message, int offset, int length) =>
            MutableDiagnostics.Add(new HtmlCssSyntaxDiagnostic(code, message, new HtmlCssSourceSpan(offset, length)));
    }

    private sealed class Parser {
        private readonly ParseContext _context;
        private readonly int _end;
        private readonly int _depth;
        private int _position;

        internal Parser(ParseContext context, int start, int end, int depth) {
            _context = context; _position = start; _end = end; _depth = depth;
        }

        internal IReadOnlyList<HtmlCssSyntaxNode> ParseContents(bool topLevel) {
            var contents = new List<HtmlCssSyntaxNode>();
            while (_position < _end) {
                _context.Cancellation.ThrowIfCancellationRequested();
                SkipIgnorable(topLevel);
                if (_position >= _end) break;
                HtmlCssSyntaxNode item;
                if (Current.Kind == HtmlCssTokenKind.AtKeyword) item = ConsumeAtRule();
                else if (topLevel) item = ConsumeQualifiedRule(stopAtSemicolon: false);
                else if (LooksLikeDeclaration()) item = ConsumeDeclaration();
                else item = ConsumeQualifiedRule(stopAtSemicolon: true);
                contents.Add(item);
            }
            return HtmlCssSyntaxCollections.ReadOnly(contents);
        }

        private HtmlCssSyntaxNode ConsumeAtRule() {
            int startIndex = _position;
            HtmlCssToken at = ConsumeToken();
            var prelude = new List<HtmlCssComponentValue>();
            while (_position < _end) {
                HtmlCssToken token = Current;
                if (token.Kind == HtmlCssTokenKind.Semicolon) {
                    _position++;
                    _context.RecordNode();
                    return new HtmlCssAtRule(_context.Source, Span(startIndex, _position), at.Value ?? string.Empty,
                        HtmlCssSyntaxCollections.ReadOnly(prelude), null, Array.Empty<HtmlCssSyntaxNode>());
                }
                if (token.Kind == HtmlCssTokenKind.OpenBrace) {
                    BlockResult result = ConsumeBlock();
                    IReadOnlyList<HtmlCssSyntaxNode> contents = new Parser(_context, result.ContentStart, result.ContentEnd, _depth + 1).ParseContents(topLevel: false);
                    _context.RecordNode();
                    return new HtmlCssAtRule(_context.Source, Span(startIndex, _position), at.Value ?? string.Empty,
                        HtmlCssSyntaxCollections.ReadOnly(prelude), result.Block, contents);
                }
                prelude.Add(ConsumeComponentValue());
            }
            _context.RecordNode();
            return new HtmlCssAtRule(_context.Source, Span(startIndex, _position), at.Value ?? string.Empty,
                HtmlCssSyntaxCollections.ReadOnly(prelude), null, Array.Empty<HtmlCssSyntaxNode>());
        }

        private HtmlCssSyntaxNode ConsumeQualifiedRule(bool stopAtSemicolon) {
            int startIndex = _position;
            var prelude = new List<HtmlCssComponentValue>();
            while (_position < _end) {
                HtmlCssToken token = Current;
                if (token.Kind == HtmlCssTokenKind.OpenBrace) {
                    BlockResult result = ConsumeBlock();
                    IReadOnlyList<HtmlCssSyntaxNode> contents = new Parser(_context, result.ContentStart, result.ContentEnd, _depth + 1).ParseContents(topLevel: false);
                    _context.RecordNode();
                    return new HtmlCssQualifiedRule(_context.Source, Span(startIndex, _position), HtmlCssSyntaxCollections.ReadOnly(prelude), result.Block, contents);
                }
                if (stopAtSemicolon && token.Kind == HtmlCssTokenKind.Semicolon) {
                    _position++;
                    return Invalid(startIndex, _position, "CSS001", "Source before the semicolon did not form a declaration or nested rule.");
                }
                prelude.Add(ConsumeComponentValue());
            }
            return Invalid(startIndex, _position, "CSS002", "A qualified rule ended before an opening curly block.");
        }

        private HtmlCssDeclaration ConsumeDeclaration() {
            int startIndex = _position;
            HtmlCssToken name = ConsumeToken();
            while (_position < _end && IsTrivia(Current.Kind)) _position++;
            HtmlCssToken colon = ConsumeToken(); // LooksLikeDeclaration already established the colon.
            int valueStart = _position < _end ? Current.Offset : colon.Offset + colon.Length;
            var values = new List<HtmlCssComponentValue>();
            while (_position < _end && Current.Kind != HtmlCssTokenKind.Semicolon) values.Add(ConsumeComponentValue());
            int valueEnd = _position < _end ? Current.Offset : EndOffsetFromPosition(_position);
            bool important = HasImportantSuffix(values);
            if (_position < _end && Current.Kind == HtmlCssTokenKind.Semicolon) _position++;
            _context.RecordNode();
            return new HtmlCssDeclaration(
                _context.Source,
                Span(startIndex, _position),
                name.Value ?? string.Empty,
                new HtmlCssSourceSpan(valueStart, Math.Max(0, valueEnd - valueStart)),
                HtmlCssSyntaxCollections.ReadOnly(values),
                important);
        }

        private BlockResult ConsumeBlock() {
            int openingIndex = _position;
            int contentStart = openingIndex + 1;
            HtmlCssSimpleBlock block = (HtmlCssSimpleBlock)ConsumeComponentValue();
            int contentEnd = block.IsClosed ? _position - 1 : _position;
            return new BlockResult(block, contentStart, contentEnd);
        }

        private HtmlCssComponentValue ConsumeComponentValue() {
            HtmlCssToken token = Current;
            if (token.Kind == HtmlCssTokenKind.Function) return ConsumeFunction();
            if (token.Kind == HtmlCssTokenKind.OpenParenthesis || token.Kind == HtmlCssTokenKind.OpenBracket || token.Kind == HtmlCssTokenKind.OpenBrace)
                return ConsumeSimpleBlock();
            _position++;
            _context.RecordNode();
            return new HtmlCssTokenValue(_context.Source, token);
        }

        private HtmlCssFunctionValue ConsumeFunction() {
            int startIndex = _position;
            HtmlCssToken opening = ConsumeToken();
            int nestedDepth = _depth + 1;
            _context.CheckDepth(nestedDepth, opening);
            var values = new List<HtmlCssComponentValue>();
            while (_position < _end && Current.Kind != HtmlCssTokenKind.CloseParenthesis) values.Add(ConsumeComponentValueAtDepth(nestedDepth));
            bool closed = _position < _end;
            if (closed) _position++;
            else _context.Diagnose("CSS003", "A function was closed by the surrounding boundary or end of input.", opening.Offset, opening.Length);
            _context.RecordNode();
            return new HtmlCssFunctionValue(_context.Source, Span(startIndex, _position), opening.Value ?? string.Empty,
                HtmlCssSyntaxCollections.ReadOnly(values), closed);
        }

        private HtmlCssSimpleBlock ConsumeSimpleBlock() {
            int startIndex = _position;
            HtmlCssToken opening = ConsumeToken();
            int nestedDepth = _depth + 1;
            _context.CheckDepth(nestedDepth, opening);
            HtmlCssTokenKind closing = opening.Kind == HtmlCssTokenKind.OpenParenthesis ? HtmlCssTokenKind.CloseParenthesis
                : opening.Kind == HtmlCssTokenKind.OpenBracket ? HtmlCssTokenKind.CloseBracket
                : HtmlCssTokenKind.CloseBrace;
            var values = new List<HtmlCssComponentValue>();
            while (_position < _end && Current.Kind != closing) values.Add(ConsumeComponentValueAtDepth(nestedDepth));
            bool closed = _position < _end;
            if (closed) _position++;
            else _context.Diagnose("CSS004", "A simple block was closed by the surrounding boundary or end of input.", opening.Offset, opening.Length);
            _context.RecordNode();
            return new HtmlCssSimpleBlock(_context.Source, Span(startIndex, _position), opening.Kind, closing,
                HtmlCssSyntaxCollections.ReadOnly(values), closed);
        }

        private HtmlCssComponentValue ConsumeComponentValueAtDepth(int depth) {
            var nested = new Parser(_context, _position, _end, depth);
            HtmlCssComponentValue result = nested.ConsumeComponentValue();
            _position = nested._position;
            return result;
        }

        private bool LooksLikeDeclaration() {
            if (Current.Kind != HtmlCssTokenKind.Identifier) return false;
            int cursor = _position + 1;
            while (cursor < _end && IsTrivia(Token(cursor).Kind)) cursor++;
            if (cursor >= _end || Token(cursor).Kind != HtmlCssTokenKind.Colon) return false;
            string name = Current.Value ?? string.Empty;
            if (name.StartsWith("--", StringComparison.Ordinal)) return true;
            // A top-level curly block after a colon is normally the nested rule's block. Generic
            // syntax cannot validate future property grammars, so the raw block remains authoritative.
            int nesting = 0;
            for (cursor++; cursor < _end; cursor++) {
                HtmlCssTokenKind kind = Token(cursor).Kind;
                if (nesting == 0 && kind == HtmlCssTokenKind.Semicolon) return true;
                if (nesting == 0 && kind == HtmlCssTokenKind.OpenBrace) return false;
                if (kind == HtmlCssTokenKind.Function || kind == HtmlCssTokenKind.OpenParenthesis || kind == HtmlCssTokenKind.OpenBracket) nesting++;
                else if (nesting > 0 && (kind == HtmlCssTokenKind.CloseParenthesis || kind == HtmlCssTokenKind.CloseBracket)) nesting--;
            }
            return true;
        }

        private static bool HasImportantSuffix(IReadOnlyList<HtmlCssComponentValue> values) {
            int index = values.Count - 1;
            while (index >= 0 && IsTrivia(values[index])) index--;
            if (index < 0 || values[index] is not HtmlCssTokenValue important
                || important.Token.Kind != HtmlCssTokenKind.Identifier
                || !string.Equals(important.Token.Value, "important", StringComparison.OrdinalIgnoreCase)) return false;
            index--;
            while (index >= 0 && IsTrivia(values[index])) index--;
            return index >= 0 && values[index] is HtmlCssTokenValue bang
                && bang.Token.Kind == HtmlCssTokenKind.Delimiter && bang.Token.Value == "!";
        }

        private static bool IsTrivia(HtmlCssComponentValue value) =>
            value is HtmlCssTokenValue token && IsTrivia(token.Token.Kind);
        private static bool IsTrivia(HtmlCssTokenKind kind) => kind == HtmlCssTokenKind.Whitespace || kind == HtmlCssTokenKind.Comment;
        private void SkipIgnorable(bool topLevel) {
            while (_position < _end) {
                HtmlCssTokenKind kind = Current.Kind;
                if (IsTrivia(kind) || kind == HtmlCssTokenKind.Semicolon || topLevel && (kind == HtmlCssTokenKind.Cdo || kind == HtmlCssTokenKind.Cdc)) _position++;
                else break;
            }
        }
        private HtmlCssInvalidSyntax Invalid(int start, int end, string code, string message) {
            HtmlCssSourceSpan span = Span(start, end);
            _context.Diagnose(code, message, span.Offset, span.Length);
            _context.RecordNode();
            return new HtmlCssInvalidSyntax(_context.Source, span, code);
        }
        private HtmlCssToken ConsumeToken() { _context.Cancellation.ThrowIfCancellationRequested(); return Token(_position++); }
        private HtmlCssToken Current => Token(_position);
        private HtmlCssToken Token(int index) => _context.Tokens[index];
        private HtmlCssSourceSpan Span(int startIndex, int endIndex) {
            int start = Token(startIndex).Offset;
            int end = EndOffsetFromPosition(endIndex);
            return new HtmlCssSourceSpan(start, Math.Max(0, end - start));
        }
        private int EndOffset(int startIndex) => Token(startIndex).Offset + Token(startIndex).Length;
        private int EndOffsetFromPosition(int position) => position <= 0 ? 0
            : position <= _end ? Token(position - 1).Offset + Token(position - 1).Length
            : Token(_end).Offset;
        private readonly struct BlockResult {
            internal BlockResult(HtmlCssSimpleBlock block, int contentStart, int contentEnd) { Block = block; ContentStart = contentStart; ContentEnd = contentEnd; }
            internal HtmlCssSimpleBlock Block { get; }
            internal int ContentStart { get; }
            internal int ContentEnd { get; }
        }
    }
}
