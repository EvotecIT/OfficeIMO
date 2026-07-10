namespace OfficeIMO.Latex;

internal sealed class LatexStructuralParser {
    private readonly LatexSourceText _source;
    private readonly IReadOnlyList<LatexToken> _tokens;
    private readonly LatexParseOptions _options;
    private readonly List<LatexDiagnostic> _diagnostics;
    private int _index;

    internal LatexStructuralParser(
        LatexSourceText source,
        IReadOnlyList<LatexToken> tokens,
        LatexParseOptions options,
        List<LatexDiagnostic> diagnostics) {
        _source = source;
        _tokens = tokens;
        _options = options;
        _diagnostics = diagnostics;
    }

    internal LatexSyntaxTree Parse() {
        var children = new List<LatexSyntaxNode>();
        while (_index < _tokens.Count) children.Add(ParseNode(0, true));
        LatexSyntaxNode root = Node(LatexSyntaxKind.Document, 0, _source.Text.Length, null, children);
        return new LatexSyntaxTree(_source, root);
    }

    private LatexSyntaxNode ParseNode(int depth, bool allowMath) {
        EnforceDepth(depth);
        LatexToken token = _tokens[_index];
        switch (token.Kind) {
            case LatexTokenKind.OpenBrace: return ParseGroup(LatexTokenKind.CloseBrace, LatexSyntaxKind.RequiredGroup, depth + 1, allowMath);
            case LatexTokenKind.OpenBracket:
            case LatexTokenKind.CloseBracket:
                _index++;
                return TokenNode(token);
            case LatexTokenKind.CloseBrace:
                _diagnostics.Add(new LatexDiagnostic("LATEX002", LatexDiagnosticSeverity.Error,
                    "Unexpected closing group delimiter was source-preserved.", token.Span));
                _index++;
                return TokenNode(token);
            case LatexTokenKind.Command:
                if (string.Equals(token.Value, "begin", StringComparison.Ordinal) && TryGetEnvironmentName(_index, out _)) {
                    return ParseEnvironment(depth + 1);
                }
                if (allowMath && (string.Equals(token.Value, "(", StringComparison.Ordinal) || string.Equals(token.Value, "[", StringComparison.Ordinal))) {
                    return ParseCommandMath(depth + 1);
                }
                return ParseCommand(depth + 1);
            case LatexTokenKind.MathShift when allowMath:
                return ParseDollarMath(depth + 1);
            default:
                _index++;
                return TokenNode(token);
        }
    }

    private LatexSyntaxNode ParseGroup(
        LatexTokenKind closingKind,
        LatexSyntaxKind groupKind,
        int depth,
        bool allowMath) {
        EnforceDepth(depth);
        LatexToken opening = _tokens[_index++];
        var children = new List<LatexSyntaxNode> {
            Node(LatexSyntaxKind.GroupDelimiter, opening.Span.Start.Offset, opening.Span.End.Offset, null)
        };
        bool terminated = false;
        while (_index < _tokens.Count) {
            LatexToken token = _tokens[_index];
            if (token.Kind == closingKind) {
                _index++;
                children.Add(Node(LatexSyntaxKind.GroupDelimiter, token.Span.Start.Offset, token.Span.End.Offset, null));
                terminated = true;
                break;
            }
            children.Add(ParseNode(depth, allowMath));
        }
        int end = terminated ? children[children.Count - 1].Span.End.Offset : _source.Text.Length;
        if (!terminated) {
            _diagnostics.Add(new LatexDiagnostic("LATEX001", LatexDiagnosticSeverity.Error,
                "Group is not terminated; source was preserved through end of input.", opening.Span));
        }
        return Node(groupKind, opening.Span.Start.Offset, end, null, children);
    }

    private LatexSyntaxNode ParseCommand(int depth, LatexCommandSyntaxSignature? explicitSignature = null) {
        EnforceDepth(depth);
        LatexToken command = _tokens[_index++];
        var children = new List<LatexSyntaxNode> {
            Node(LatexSyntaxKind.CommandToken, command.Span.Start.Offset, command.Span.End.Offset, command.Value)
        };
        int end = command.Span.End.Offset;
        LatexCommandSyntaxSignature? signature = explicitSignature ?? LatexProfileSyntaxCatalog.GetCommand(command.Value ?? string.Empty);
        if (signature != null) {
            if (signature.AllowsStar && _index < _tokens.Count &&
                _tokens[_index].Kind == LatexTokenKind.Text && string.Equals(_tokens[_index].Text, "*", StringComparison.Ordinal)) {
                LatexToken star = _tokens[_index++];
                children.Add(TokenNode(star));
                end = star.Span.End.Offset;
            }
            for (int index = 0; index < signature.Arguments.Count; index++) {
                LatexTokenKind openingKind = signature.Arguments[index] == LatexArgumentGroupKind.Optional
                    ? LatexTokenKind.OpenBracket
                    : LatexTokenKind.OpenBrace;
                if (!TryParseCommandGroup(openingKind, depth, children, ref end)) {
                    if (signature.Arguments[index] == LatexArgumentGroupKind.Required) break;
                }
            }
            return Node(LatexSyntaxKind.Command, command.Span.Start.Offset, end, command.Value, children);
        }

        ParseFallbackCommandGroups(depth, children, ref end);
        return Node(LatexSyntaxKind.Command, command.Span.Start.Offset, end, command.Value, children);
    }

    private void ParseFallbackCommandGroups(int depth, List<LatexSyntaxNode> children, ref int end) {
        while (_index < _tokens.Count) {
            int lookahead = _index;
            while (lookahead < _tokens.Count && IsArgumentTrivia(_tokens[lookahead])) lookahead++;
            if (lookahead >= _tokens.Count ||
                (_tokens[lookahead].Kind != LatexTokenKind.OpenBrace && _tokens[lookahead].Kind != LatexTokenKind.OpenBracket)) break;
            while (_index < lookahead) {
                LatexToken trivia = _tokens[_index++];
                children.Add(TokenNode(trivia));
                end = trivia.Span.End.Offset;
            }
            LatexToken opening = _tokens[_index];
            LatexSyntaxNode group = opening.Kind == LatexTokenKind.OpenBrace
                ? ParseGroup(LatexTokenKind.CloseBrace, LatexSyntaxKind.RequiredGroup, depth + 1, true)
                : ParseGroup(LatexTokenKind.CloseBracket, LatexSyntaxKind.OptionalGroup, depth + 1, true);
            children.Add(group);
            end = group.Span.End.Offset;
        }
    }

    private bool TryParseCommandGroup(
        LatexTokenKind expectedOpening,
        int depth,
        List<LatexSyntaxNode> children,
        ref int end) {
        int lookahead = _index;
        while (lookahead < _tokens.Count && IsArgumentTrivia(_tokens[lookahead])) lookahead++;
        if (lookahead >= _tokens.Count || _tokens[lookahead].Kind != expectedOpening) return false;
        while (_index < lookahead) {
            LatexToken trivia = _tokens[_index++];
            children.Add(TokenNode(trivia));
            end = trivia.Span.End.Offset;
        }
        LatexSyntaxNode group = expectedOpening == LatexTokenKind.OpenBrace
            ? ParseGroup(LatexTokenKind.CloseBrace, LatexSyntaxKind.RequiredGroup, depth + 1, true)
            : ParseGroup(LatexTokenKind.CloseBracket, LatexSyntaxKind.OptionalGroup, depth + 1, true);
        children.Add(group);
        end = group.Span.End.Offset;
        return true;
    }

    private LatexSyntaxNode ParseEnvironment(int depth) {
        EnforceDepth(depth);
        int start = _tokens[_index].Span.Start.Offset;
        TryGetEnvironmentName(_index, out string expectedName);
        LatexSyntaxNode begin = ParseCommand(depth + 1, LatexProfileSyntaxCatalog.GetEnvironmentBegin(expectedName));
        string name = GetFirstRequiredGroupContent(begin) ?? string.Empty;
        var children = new List<LatexSyntaxNode> { begin };
        LatexSyntaxNode? endCommand = null;
        while (_index < _tokens.Count) {
            if (_tokens[_index].Kind == LatexTokenKind.Command &&
                string.Equals(_tokens[_index].Value, "end", StringComparison.Ordinal) &&
                TryGetEnvironmentName(_index, out string endName)) {
                LatexSyntaxNode candidate = ParseCommand(depth + 1, LatexProfileSyntaxCatalog.EnvironmentEnd);
                children.Add(candidate);
                if (string.Equals(name, endName, StringComparison.Ordinal)) {
                    endCommand = candidate;
                    break;
                }
                _diagnostics.Add(new LatexDiagnostic("LATEX005", LatexDiagnosticSeverity.Error,
                    "Environment '" + name + "' encountered mismatched end '" + endName + "'.", candidate.Span));
                continue;
            }
            children.Add(ParseNode(depth + 1, true));
        }
        int end = endCommand?.Span.End.Offset ?? _source.Text.Length;
        if (endCommand == null) {
            _diagnostics.Add(new LatexDiagnostic("LATEX004", LatexDiagnosticSeverity.Error,
                "Environment '" + name + "' is not terminated.", begin.Span));
        }
        return Node(LatexSyntaxKind.Environment, start, end, name, children);
    }

    private LatexSyntaxNode ParseDollarMath(int depth) {
        EnforceDepth(depth);
        LatexToken opening = _tokens[_index++];
        var children = new List<LatexSyntaxNode> {
            Node(LatexSyntaxKind.MathDelimiter, opening.Span.Start.Offset, opening.Span.End.Offset, opening.Text)
        };
        bool terminated = false;
        while (_index < _tokens.Count) {
            LatexToken token = _tokens[_index];
            if (token.Kind == LatexTokenKind.MathShift && string.Equals(token.Text, opening.Text, StringComparison.Ordinal)) {
                _index++;
                children.Add(Node(LatexSyntaxKind.MathDelimiter, token.Span.Start.Offset, token.Span.End.Offset, token.Text));
                terminated = true;
                break;
            }
            children.Add(ParseNode(depth + 1, false));
        }
        int end = terminated ? children[children.Count - 1].Span.End.Offset : _source.Text.Length;
        if (!terminated) {
            _diagnostics.Add(new LatexDiagnostic("LATEX003", LatexDiagnosticSeverity.Error,
                "Math region is not terminated.", opening.Span));
        }
        return Node(LatexSyntaxKind.Math, opening.Span.Start.Offset, end, opening.Text, children);
    }

    private LatexSyntaxNode ParseCommandMath(int depth) {
        EnforceDepth(depth);
        LatexToken opening = _tokens[_index++];
        string closingName = string.Equals(opening.Value, "(", StringComparison.Ordinal) ? ")" : "]";
        var children = new List<LatexSyntaxNode> {
            Node(LatexSyntaxKind.MathDelimiter, opening.Span.Start.Offset, opening.Span.End.Offset, opening.Text)
        };
        bool terminated = false;
        while (_index < _tokens.Count) {
            LatexToken token = _tokens[_index];
            if (token.Kind == LatexTokenKind.Command && string.Equals(token.Value, closingName, StringComparison.Ordinal)) {
                _index++;
                children.Add(Node(LatexSyntaxKind.MathDelimiter, token.Span.Start.Offset, token.Span.End.Offset, token.Text));
                terminated = true;
                break;
            }
            children.Add(ParseNode(depth + 1, false));
        }
        int end = terminated ? children[children.Count - 1].Span.End.Offset : _source.Text.Length;
        if (!terminated) {
            _diagnostics.Add(new LatexDiagnostic("LATEX003", LatexDiagnosticSeverity.Error,
                "Math region is not terminated.", opening.Span));
        }
        return Node(LatexSyntaxKind.Math, opening.Span.Start.Offset, end, opening.Text, children);
    }

    private bool TryGetEnvironmentName(int commandIndex, out string name) {
        name = string.Empty;
        int index = commandIndex + 1;
        while (index < _tokens.Count && IsArgumentTrivia(_tokens[index])) index++;
        if (index >= _tokens.Count || _tokens[index].Kind != LatexTokenKind.OpenBrace) return false;
        int depth = 1;
        int contentStart = _tokens[index].Span.End.Offset;
        for (int current = index + 1; current < _tokens.Count; current++) {
            if (_tokens[current].Kind == LatexTokenKind.OpenBrace) depth++;
            else if (_tokens[current].Kind == LatexTokenKind.CloseBrace && --depth == 0) {
                name = _source.Text.Substring(contentStart, _tokens[current].Span.Start.Offset - contentStart).Trim();
                return name.Length > 0;
            }
        }
        return false;
    }

    private string? GetFirstRequiredGroupContent(LatexSyntaxNode command) {
        LatexSyntaxNode? group = command.Children.FirstOrDefault(static child => child.Kind == LatexSyntaxKind.RequiredGroup);
        if (group == null || group.Children.Count < 2) return null;
        int start = group.Children[0].Span.End.Offset;
        int end = group.Children[group.Children.Count - 1].Kind == LatexSyntaxKind.GroupDelimiter
            ? group.Children[group.Children.Count - 1].Span.Start.Offset
            : group.Span.End.Offset;
        return _source.Text.Substring(start, end - start).Trim();
    }

    private LatexSyntaxNode TokenNode(LatexToken token) {
        LatexSyntaxKind kind = token.Kind == LatexTokenKind.Comment
            ? LatexSyntaxKind.Comment
            : token.Kind == LatexTokenKind.Whitespace || token.Kind == LatexTokenKind.LineEnding
                ? LatexSyntaxKind.Trivia
                : token.Kind == LatexTokenKind.Command
                    ? LatexSyntaxKind.CommandToken
                    : LatexSyntaxKind.Text;
        return Node(kind, token.Span.Start.Offset, token.Span.End.Offset, token.Value);
    }

    private LatexSyntaxNode Node(
        LatexSyntaxKind kind,
        int start,
        int end,
        string? value,
        IReadOnlyList<LatexSyntaxNode>? children = null) {
        IReadOnlyList<LatexSyntaxNode>? completed = CompleteCoverage(start, end, children);
        return new LatexSyntaxNode(kind, _source.CreateSpan(start, end), _source.Text.Substring(start, end - start), value, completed);
    }

    private IReadOnlyList<LatexSyntaxNode>? CompleteCoverage(int start, int end, IReadOnlyList<LatexSyntaxNode>? children) {
        if (children == null || children.Count == 0) return children;
        var result = new List<LatexSyntaxNode>(children.Count + 2);
        int expected = start;
        for (int index = 0; index < children.Count; index++) {
            LatexSyntaxNode child = children[index];
            if (child.Span.Start.Offset > expected) result.Add(Node(LatexSyntaxKind.Trivia, expected, child.Span.Start.Offset, null));
            result.Add(child);
            expected = child.Span.End.Offset;
        }
        if (expected < end) result.Add(Node(LatexSyntaxKind.Trivia, expected, end, null));
        return result;
    }

    private void EnforceDepth(int depth) {
        if (depth > _options.MaximumNestingDepth) throw new InvalidDataException("LaTeX source exceeds MaximumNestingDepth.");
    }

    private static bool IsArgumentTrivia(LatexToken token) =>
        token.Kind == LatexTokenKind.Whitespace || token.Kind == LatexTokenKind.LineEnding || token.Kind == LatexTokenKind.Comment;
}
