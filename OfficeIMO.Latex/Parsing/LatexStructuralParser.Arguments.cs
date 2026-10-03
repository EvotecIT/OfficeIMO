namespace OfficeIMO.Latex;

internal sealed partial class LatexStructuralParser {
    private void TryParseStarModifier(List<LatexSyntaxNode> children, ref int end) {
        int lookahead = FindCommandArgumentStart();
        if (lookahead >= _tokens.Count || _tokens[lookahead].Kind != LatexTokenKind.Text) return;
        LatexTokenView token = _tokens[lookahead];
        int start = token.StartOffset + _textTokenOffset;
        if (_source.Text[start] != '*') return;
        while (_index < lookahead) children.Add(TokenNode(_tokens[_index++]));
        children.Add(Node(LatexSyntaxKind.Text, start, start + 1, null));
        end = start + 1;
        if (end == token.EndOffset) {
            _index++;
            _textTokenOffset = 0;
        } else {
            _textTokenOffset = end - token.StartOffset;
        }
    }

    // The tokenizer deliberately coalesces ordinary text. Argument binding consumes
    // one Unicode scalar without splitting or changing the public token inventory.
    private bool TryParseSingleTokenArgument(int depth, List<LatexSyntaxNode> children, ref int end) {
        EnforceDepth(depth + 1);
        int lookahead = FindCommandArgumentStart();
        if (lookahead >= _tokens.Count) return false;
        LatexTokenView token = _tokens[lookahead];
        if (token.Kind == LatexTokenKind.CloseBrace || token.Kind == LatexTokenKind.OpenBrace ||
            token.Kind == LatexTokenKind.Verbatim || token.Kind == LatexTokenKind.MathShift ||
            IsArgumentTrivia(token)) return false;
        while (_index < lookahead) children.Add(TokenNode(_tokens[_index++]));

        int start = token.StartOffset + _textTokenOffset;
        int tokenEnd = token.EndOffset;
        LatexSyntaxNode value;
        if (token.Kind == LatexTokenKind.Command) {
            var control = Node(LatexSyntaxKind.CommandToken, start, tokenEnd, token.Value);
            value = Node(LatexSyntaxKind.Command, start, tokenEnd, token.Value, new[] { control });
            _index++;
        } else {
            tokenEnd = start + 1;
            if (char.IsHighSurrogate(_source.Text[start]) && tokenEnd < token.EndOffset &&
                char.IsLowSurrogate(_source.Text[tokenEnd])) tokenEnd++;
            value = Node(LatexSyntaxKind.Text, start, tokenEnd, null);
            if (tokenEnd == token.EndOffset) {
                _index++;
                _textTokenOffset = 0;
            } else {
                _textTokenOffset = tokenEnd - token.StartOffset;
            }
        }
        children.Add(Node(LatexSyntaxKind.SingleTokenArgument, start, tokenEnd, null, new[] { value }));
        end = tokenEnd;

        // A captured control word still discards its lexical delimiter whitespace.
        // Keep that source outside the argument so edits cannot duplicate or erase it.
        if (token.Kind == LatexTokenKind.Command && token.Value?.Length > 0 &&
            (char.IsLetter(token.Value[0]) || token.Value[0] == '@')) {
            int delimiterEnd = LatexWhitespaceSyntax.SkipControlWordDelimiter(
                _source.Text, tokenEnd, _source.Text.Length, _cancellationToken);
            while (_index < _tokens.Count && _tokens[_index].EndOffset <= delimiterEnd) {
                _cancellationToken.ThrowIfCancellationRequested();
                children.Add(TokenNode(_tokens[_index++]));
            }
            end = delimiterEnd;
        }
        return true;
    }

    private int FindCommandArgumentStart() {
        if (_index >= _tokens.Count) return _index;
        int delimiterEnd = LatexWhitespaceSyntax.SkipControlWordDelimiter(_source.Text,
            _tokens[_index].StartOffset + _textTokenOffset, _source.Text.Length, _cancellationToken);
        int lookahead = _index;
        while (lookahead < _tokens.Count && IsArgumentTrivia(_tokens[lookahead]) &&
            _tokens[lookahead].EndOffset <= delimiterEnd) {
            _cancellationToken.ThrowIfCancellationRequested();
            lookahead++;
        }
        return lookahead;
    }
}
