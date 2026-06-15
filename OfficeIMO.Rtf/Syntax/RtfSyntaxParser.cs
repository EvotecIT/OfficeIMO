using OfficeIMO.Rtf.Diagnostics;

namespace OfficeIMO.Rtf.Syntax;

/// <summary>
/// Parses RTF tokens into a nested syntax tree.
/// </summary>
public static class RtfSyntaxParser {
    /// <summary>
    /// Parses RTF text into a syntax tree.
    /// </summary>
    public static RtfSyntaxTree Parse(string rtf) {
        if (rtf == null) throw new ArgumentNullException(nameof(rtf));
        RtfTokenizeResult tokenized = RtfTokenizer.Tokenize(rtf);
        var diagnostics = new List<RtfDiagnostic>(tokenized.Diagnostics);
        var parser = new Parser(tokenized.Tokens, diagnostics);
        return parser.Parse();
    }

    private sealed class Parser {
        private readonly IReadOnlyList<RtfToken> _tokens;
        private readonly List<RtfDiagnostic> _diagnostics;
        private int _index;

        public Parser(IReadOnlyList<RtfToken> tokens, List<RtfDiagnostic> diagnostics) {
            _tokens = tokens;
            _diagnostics = diagnostics;
        }

        public RtfSyntaxTree Parse() {
            if (Current.Kind == RtfTokenKind.GroupStart) {
                RtfGroup root = ParseGroup();
                while (Current.Kind != RtfTokenKind.EndOfFile) {
                    if (Current.Kind == RtfTokenKind.GroupEnd) {
                        _diagnostics.Add(new RtfDiagnostic(RtfDiagnosticSeverity.Warning, "RTF010", "Ignoring unmatched closing brace after the root group.", Current.Position));
                    } else {
                        _diagnostics.Add(new RtfDiagnostic(RtfDiagnosticSeverity.Warning, "RTF011", "Ignoring content after the root group.", Current.Position));
                    }

                    _index++;
                }

                return new RtfSyntaxTree(root, _diagnostics.AsReadOnly());
            }

            _diagnostics.Add(new RtfDiagnostic(RtfDiagnosticSeverity.Warning, "RTF012", "RTF input does not start with a group; a synthetic root group was created.", Current.Position));
            var children = new List<RtfNode>();
            while (Current.Kind != RtfTokenKind.EndOfFile) {
                children.Add(ParseNode());
            }

            return new RtfSyntaxTree(new RtfGroup(0, children), _diagnostics.AsReadOnly());
        }

        private RtfGroup ParseGroup() {
            int position = Current.Position;
            Expect(RtfTokenKind.GroupStart);
            var children = new List<RtfNode>();
            while (Current.Kind != RtfTokenKind.EndOfFile && Current.Kind != RtfTokenKind.GroupEnd) {
                children.Add(ParseNode());
            }

            if (Current.Kind == RtfTokenKind.GroupEnd) {
                _index++;
            } else {
                _diagnostics.Add(new RtfDiagnostic(RtfDiagnosticSeverity.Error, "RTF013", "RTF group was not closed before end of input.", position));
            }

            return new RtfGroup(position, children);
        }

        private RtfNode ParseNode() {
            RtfToken token = Current;
            switch (token.Kind) {
                case RtfTokenKind.GroupStart:
                    return ParseGroup();
                case RtfTokenKind.GroupEnd:
                    _diagnostics.Add(new RtfDiagnostic(RtfDiagnosticSeverity.Warning, "RTF014", "Ignoring unmatched closing brace.", token.Position));
                    _index++;
                    return new RtfText(token.Position, string.Empty, string.Empty);
                case RtfTokenKind.ControlWord:
                    _index++;
                    return new RtfControlWord(token.Position, token.ControlName ?? string.Empty, token.Parameter, token.HasParameter, token.RawText);
                case RtfTokenKind.ControlSymbol:
                    _index++;
                    return new RtfControlSymbol(token.Position, token.ControlSymbol ?? '\0', token.Parameter, token.HasParameter, token.RawText);
                case RtfTokenKind.Text:
                    _index++;
                    return new RtfText(token.Position, token.Text ?? string.Empty, token.RawText);
                case RtfTokenKind.Binary:
                    _index++;
                    return new RtfBinary(token.Position, token.BinaryData ?? Array.Empty<byte>(), token.RawText);
                default:
                    _index++;
                    return new RtfText(token.Position, string.Empty, string.Empty);
            }
        }

        private RtfToken Current => _index < _tokens.Count ? _tokens[_index] : _tokens[_tokens.Count - 1];

        private void Expect(RtfTokenKind kind) {
            if (Current.Kind != kind) {
                _diagnostics.Add(new RtfDiagnostic(RtfDiagnosticSeverity.Error, "RTF015", $"Expected token '{kind}' but found '{Current.Kind}'.", Current.Position));
                return;
            }

            _index++;
        }
    }
}
