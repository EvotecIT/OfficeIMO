namespace OfficeIMO.Latex;

/// <summary>Document-local simple macro definition.</summary>
public sealed class LatexMacroDefinition {
    internal LatexMacroDefinition(
        LatexCommand command,
        string name,
        int parameterCount,
        string? defaultValue,
        string body,
        bool isSafe) {
        Command = command;
        Name = name;
        ParameterCount = parameterCount;
        DefaultValue = defaultValue;
        Body = body;
        IsSafe = isSafe;
    }

    /// <summary>Backing new/renew/provide command.</summary>
    public LatexCommand Command { get; }
    /// <summary>Defined control sequence without backslash.</summary>
    public string Name { get; }
    /// <summary>Required parameter count, including an optional first parameter when defaulted.</summary>
    public int ParameterCount { get; }
    /// <summary>Optional first-parameter default.</summary>
    public string? DefaultValue { get; }
    /// <summary>Unexpanded replacement source.</summary>
    public string Body { get; }
    /// <summary>True when the body passes the OfficeIMO structural safety policy.</summary>
    public bool IsSafe { get; }
}

/// <summary>Diagnostic from explicit simple macro expansion.</summary>
public sealed class LatexMacroExpansionDiagnostic {
    internal LatexMacroExpansionDiagnostic(string code, LatexDiagnosticSeverity severity, string message, int offset) {
        Code = code;
        Severity = severity;
        Message = message;
        Offset = offset;
    }
    /// <summary>Stable code.</summary>
    public string Code { get; }
    /// <summary>Severity.</summary>
    public LatexDiagnosticSeverity Severity { get; }
    /// <summary>Message.</summary>
    public string Message { get; }
    /// <summary>Input offset.</summary>
    public int Offset { get; }
}

/// <summary>Explicit expansion result.</summary>
public sealed class LatexMacroExpansionResult {
    internal LatexMacroExpansionResult(string value, IReadOnlyList<LatexMacroExpansionDiagnostic> diagnostics) {
        Value = value;
        Diagnostics = diagnostics;
    }
    /// <summary>Expanded source.</summary>
    public string Value { get; }
    /// <summary>Skipped, cyclic, or limited expansion diagnostics.</summary>
    public IReadOnlyList<LatexMacroExpansionDiagnostic> Diagnostics { get; }
}

/// <summary>Bounded expander for safe document-local simple macros only.</summary>
public static class LatexSimpleMacroExpander {
    private const int DefaultMaximumTokenCount = 2_000_000;

    /// <summary>Expands safe definitions in an explicit input string.</summary>
    public static LatexMacroExpansionResult Expand(
        string value,
        IReadOnlyList<LatexMacroDefinition> definitions,
        int maximumDepth = 16,
        int maximumOutputLength = 16 * 1024 * 1024) =>
        Expand(value, definitions, maximumDepth, maximumOutputLength, 64 * 1024 * 1024);

    /// <summary>Expands safe definitions with independent recursion, output, and input limits.</summary>
    public static LatexMacroExpansionResult Expand(
        string value,
        IReadOnlyList<LatexMacroDefinition> definitions,
        int maximumDepth,
        int maximumOutputLength,
        int maximumInputLength = 64 * 1024 * 1024) {
        return Expand(value, definitions, maximumDepth, maximumOutputLength, maximumInputLength,
            DefaultMaximumTokenCount);
    }

    /// <summary>Expands safe definitions with independent recursion, output, input, and aggregate token limits.</summary>
    public static LatexMacroExpansionResult Expand(
        string value,
        IReadOnlyList<LatexMacroDefinition> definitions,
        int maximumDepth,
        int maximumOutputLength,
        int maximumInputLength,
        int maximumTokenCount) {
        if (value == null) throw new ArgumentNullException(nameof(value));
        if (definitions == null) throw new ArgumentNullException(nameof(definitions));
        if (maximumDepth < 1) throw new ArgumentOutOfRangeException(nameof(maximumDepth));
        if (maximumOutputLength < 1) throw new ArgumentOutOfRangeException(nameof(maximumOutputLength));
        if (maximumInputLength < 1) throw new ArgumentOutOfRangeException(nameof(maximumInputLength));
        if (maximumTokenCount < 1) throw new ArgumentOutOfRangeException(nameof(maximumTokenCount));
        if (value.Length > maximumInputLength) {
            throw new ArgumentException("Simple macro expansion input exceeds maximumInputLength.", nameof(value));
        }
        var map = new Dictionary<string, LatexMacroDefinition>(StringComparer.Ordinal);
        foreach (LatexMacroDefinition definition in definitions.Where(static definition => definition.IsSafe)) {
            if (string.Equals(definition.Command.Name, "providecommand", StringComparison.Ordinal) && map.ContainsKey(definition.Name)) continue;
            map[definition.Name] = definition;
        }
        var diagnostics = new List<LatexMacroExpansionDiagnostic>();
        var tokenBudget = new TokenBudget(maximumTokenCount);
        string output = ExpandCore(value, map, diagnostics, new HashSet<string>(StringComparer.Ordinal), 0,
            maximumDepth, maximumOutputLength, maximumInputLength, tokenBudget);
        return new LatexMacroExpansionResult(output, diagnostics);
    }

    private static string ExpandCore(
        string value,
        IReadOnlyDictionary<string, LatexMacroDefinition> definitions,
        List<LatexMacroExpansionDiagnostic> diagnostics,
        HashSet<string> active,
        int depth,
        int maximumDepth,
        int maximumOutputLength,
        int maximumInputLength,
        TokenBudget tokenBudget) {
        if (depth > maximumDepth) throw new InvalidDataException("Simple macro expansion exceeds maximumDepth.");
        if (value.Length == 0) return string.Empty;
        if (depth > 0 && value.Length > maximumOutputLength) {
            throw new InvalidDataException("Simple macro expansion exceeds maximumOutputLength.");
        }
        var output = new StringBuilder(value.Length);
        IReadOnlyList<LatexToken> tokens = LatexTokenizer.Tokenize(value, new LatexParseOptions {
            MaximumInputLength = depth == 0 ? maximumInputLength : maximumOutputLength,
            MaximumTokenCount = tokenBudget.GetTokenizerLimit()
        });
        tokenBudget.Consume(tokens.Count);
        for (int tokenIndex = 0; tokenIndex < tokens.Count;) {
            LatexToken invocation = tokens[tokenIndex];
            string name = invocation.Value ?? string.Empty;
            if (invocation.Kind != LatexTokenKind.Command || !definitions.TryGetValue(name, out LatexMacroDefinition? definition)) {
                output.Append(invocation.Text);
                tokenIndex++;
                EnforceLength(output, maximumOutputLength);
                continue;
            }
            if (!active.Add(name)) {
                diagnostics.Add(new LatexMacroExpansionDiagnostic("LATEXMAC002", LatexDiagnosticSeverity.Error,
                    "Cyclic simple macro invocation '" + name + "'.", invocation.Span.Start.Offset));
                output.Append(invocation.Text);
                tokenIndex++;
                continue;
            }

            int cursor = tokenIndex + 1;
            var arguments = new List<string>();
            if (definition.DefaultValue != null) {
                SkipArgumentTrivia(tokens, ref cursor);
                if (TryReadBalanced(tokens, value, ref cursor, LatexTokenKind.OpenBracket, LatexTokenKind.CloseBracket,
                        out string optional)) {
                    arguments.Add(optional);
                } else {
                    arguments.Add(definition.DefaultValue);
                }
            }
            while (arguments.Count < definition.ParameterCount) {
                SkipArgumentTrivia(tokens, ref cursor);
                if (!TryReadBalanced(tokens, value, ref cursor, LatexTokenKind.OpenBrace, LatexTokenKind.CloseBrace,
                        out string argument)) break;
                arguments.Add(argument);
            }
            if (arguments.Count != definition.ParameterCount) {
                diagnostics.Add(new LatexMacroExpansionDiagnostic("LATEXMAC001", LatexDiagnosticSeverity.Warning,
                    "Simple macro '" + name + "' did not receive the required arguments.", invocation.Span.Start.Offset));
                output.Append(invocation.Text);
                active.Remove(name);
                tokenIndex++;
                continue;
            }
            string replacement = SubstituteParameters(definition.Body, arguments, tokenBudget);
            output.Append(ExpandCore(replacement, definitions, diagnostics, active, depth + 1,
                maximumDepth, maximumOutputLength, maximumInputLength, tokenBudget));
            active.Remove(name);
            tokenIndex = cursor;
            EnforceLength(output, maximumOutputLength);
        }
        return output.ToString();
    }

    private static string SubstituteParameters(
        string body,
        IReadOnlyList<string> arguments,
        TokenBudget tokenBudget) {
        if (body.Length == 0) return string.Empty;
        var output = new StringBuilder(body.Length);
        IReadOnlyList<LatexToken> tokens = LatexTokenizer.Tokenize(body, new LatexParseOptions {
            MaximumInputLength = body.Length,
            MaximumTokenCount = tokenBudget.GetTokenizerLimit()
        });
        tokenBudget.Consume(tokens.Count);
        for (int index = 0; index < tokens.Count; index++) {
            LatexToken token = tokens[index];
            if (token.Kind == LatexTokenKind.Parameter && index + 1 < tokens.Count) {
                LatexToken next = tokens[index + 1];
                if (next.Kind == LatexTokenKind.Text && next.Text.Length > 0 && next.Text[0] >= '1' && next.Text[0] <= '9') {
                    int parameter = next.Text[0] - '1';
                    if (parameter < arguments.Count) output.Append(arguments[parameter]);
                    if (next.Text.Length > 1) output.Append(next.Text, 1, next.Text.Length - 1);
                    index++;
                    continue;
                }
            }
            output.Append(token.Text);
        }
        return output.ToString();
    }

    private static bool TryReadBalanced(
        IReadOnlyList<LatexToken> tokens,
        string value,
        ref int cursor,
        LatexTokenKind open,
        LatexTokenKind close,
        out string content) {
        content = string.Empty;
        if (cursor >= tokens.Count || tokens[cursor].Kind != open) return false;
        int start = tokens[cursor].Span.End.Offset;
        cursor++;
        int depth = 1;
        while (cursor < tokens.Count) {
            LatexToken token = tokens[cursor];
            if (token.Kind == open) depth++;
            else if (token.Kind == close && --depth == 0) {
                content = value.Substring(start, token.Span.Start.Offset - start);
                cursor++;
                return true;
            }
            cursor++;
        }
        return false;
    }

    private static void SkipArgumentTrivia(IReadOnlyList<LatexToken> tokens, ref int cursor) {
        while (cursor < tokens.Count && (tokens[cursor].Kind == LatexTokenKind.Whitespace ||
               tokens[cursor].Kind == LatexTokenKind.LineEnding || tokens[cursor].Kind == LatexTokenKind.Comment)) cursor++;
    }

    private static void EnforceLength(StringBuilder output, int maximumOutputLength) {
        if (output.Length > maximumOutputLength) throw new InvalidDataException("Simple macro expansion exceeds maximumOutputLength.");
    }

    private sealed class TokenBudget {
        internal TokenBudget(int maximumTokenCount) => Remaining = maximumTokenCount;

        internal int Remaining { get; private set; }

        internal int GetTokenizerLimit() {
            if (Remaining < 1) throw new InvalidDataException("Simple macro expansion exceeds maximumTokenCount.");
            return Remaining;
        }

        internal void Consume(int count) {
            if (count > Remaining) throw new InvalidDataException("Simple macro expansion exceeds maximumTokenCount.");
            Remaining -= count;
        }
    }
}