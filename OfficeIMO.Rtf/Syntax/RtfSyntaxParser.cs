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
        return Parse(rtf, RtfReadOptions.CreateOfficeIMOProfile(), CancellationToken.None);
    }

    /// <summary>
    /// Parses RTF text into a syntax tree while limiting nested group depth.
    /// </summary>
    public static RtfSyntaxTree Parse(string rtf, int maxDepth) {
        return Parse(rtf, new RtfReadOptions { MaxDepth = maxDepth }, CancellationToken.None);
    }

    /// <summary>
    /// Parses RTF text using the configured resource limits and cancellation token.
    /// </summary>
    public static RtfSyntaxTree Parse(string rtf, RtfReadOptions? options, CancellationToken cancellationToken = default) {
        if (rtf == null) throw new ArgumentNullException(nameof(rtf));
        RtfReadOptions readOptions = options ?? RtfReadOptions.CreateOfficeIMOProfile();
        RtfTokenizeResult tokenized = RtfTokenizer.Tokenize(rtf, readOptions, cancellationToken);
        var diagnostics = new List<RtfDiagnostic>(tokenized.Diagnostics);
        var parser = new Parser(rtf, tokenized.Tokens, diagnostics, readOptions.MaxDepth, cancellationToken);
        return parser.Parse();
    }

    private sealed class Parser {
        private readonly IReadOnlyList<RtfToken> _tokens;
        private readonly string _source;
        private readonly List<RtfDiagnostic> _diagnostics;
        private readonly int _maxDepth;
        private readonly CancellationToken _cancellationToken;
        private int _index;
        private int _operationCount;

        public Parser(string source, IReadOnlyList<RtfToken> tokens, List<RtfDiagnostic> diagnostics, int maxDepth, CancellationToken cancellationToken) {
            _source = source;
            _tokens = tokens;
            _diagnostics = diagnostics;
            _maxDepth = maxDepth;
            _cancellationToken = cancellationToken;
        }

        public RtfSyntaxTree Parse() {
            if (Current.Kind == RtfTokenKind.GroupStart) {
                RtfGroup root = ParseGroup(depth: 0);
                int suffixStart = Current.Position;
                while (Current.Kind != RtfTokenKind.EndOfFile) {
                    if (Current.Kind == RtfTokenKind.GroupEnd) {
                        _diagnostics.Add(new RtfDiagnostic(RtfDiagnosticSeverity.Warning, "RTF010", "Ignoring unmatched closing brace after the root group.", Current.Position));
                    } else if (Current.Kind != RtfTokenKind.Text || !string.IsNullOrWhiteSpace(Current.Text)) {
                        _diagnostics.Add(new RtfDiagnostic(RtfDiagnosticSeverity.Warning, "RTF011", "Ignoring content after the root group.", Current.Position));
                    }

                    _index++;
                }

                string prefix = root.Position > 0 ? _source.Substring(0, root.Position) : string.Empty;
                string suffix = suffixStart < _source.Length ? _source.Substring(suffixStart) : string.Empty;
                return new RtfSyntaxTree(root, _diagnostics.AsReadOnly(), prefix, suffix);
            }

            _diagnostics.Add(new RtfDiagnostic(RtfDiagnosticSeverity.Warning, "RTF012", "RTF input does not start with a group; a synthetic root group was created.", Current.Position));
            var children = new List<RtfNode>();
            while (Current.Kind != RtfTokenKind.EndOfFile) {
                children.Add(ParseNode(depth: 0));
            }

            return new RtfSyntaxTree(new RtfGroup(0, children), _diagnostics.AsReadOnly());
        }

        private RtfGroup ParseGroup(int depth) {
            CheckCancellation();
            int position = Current.Position;
            Expect(RtfTokenKind.GroupStart);
            if (depth > _maxDepth) {
                _diagnostics.Add(new RtfDiagnostic(RtfDiagnosticSeverity.Error, "RTF100", "Maximum RTF group depth was exceeded.", position));
                SkipCurrentGroup(position);
                return new RtfGroup(position, Array.Empty<RtfNode>());
            }

            var children = new List<RtfNode>();
            while (Current.Kind != RtfTokenKind.EndOfFile && Current.Kind != RtfTokenKind.GroupEnd) {
                children.Add(ParseNode(depth));
            }

            if (Current.Kind == RtfTokenKind.GroupEnd) {
                _index++;
            } else {
                _diagnostics.Add(new RtfDiagnostic(RtfDiagnosticSeverity.Error, "RTF013", "RTF group was not closed before end of input.", position));
            }

            return new RtfGroup(position, children);
        }

        private RtfNode ParseNode(int depth) {
            CheckCancellation();
            RtfToken token = Current;
            switch (token.Kind) {
                case RtfTokenKind.GroupStart:
                    return ParseGroup(depth + 1);
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

        private void SkipCurrentGroup(int position) {
            int nestedGroups = 1;
            while (Current.Kind != RtfTokenKind.EndOfFile && nestedGroups > 0) {
                CheckCancellation();
                if (Current.Kind == RtfTokenKind.GroupStart) {
                    nestedGroups++;
                } else if (Current.Kind == RtfTokenKind.GroupEnd) {
                    nestedGroups--;
                }

                _index++;
            }

            if (nestedGroups > 0) {
                _diagnostics.Add(new RtfDiagnostic(RtfDiagnosticSeverity.Error, "RTF013", "RTF group was not closed before end of input.", position));
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

        private void CheckCancellation() {
            if ((_operationCount++ & 0x3FF) == 0) {
                _cancellationToken.ThrowIfCancellationRequested();
            }
        }
    }
}
