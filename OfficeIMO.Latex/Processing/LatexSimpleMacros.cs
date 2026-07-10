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
    /// <summary>Expands safe definitions in an explicit input string.</summary>
    public static LatexMacroExpansionResult Expand(
        string value,
        IReadOnlyList<LatexMacroDefinition> definitions,
        int maximumDepth = 16,
        int maximumOutputLength = 16 * 1024 * 1024) {
        if (value == null) throw new ArgumentNullException(nameof(value));
        if (definitions == null) throw new ArgumentNullException(nameof(definitions));
        if (maximumDepth < 1) throw new ArgumentOutOfRangeException(nameof(maximumDepth));
        if (maximumOutputLength < 1) throw new ArgumentOutOfRangeException(nameof(maximumOutputLength));
        var map = new Dictionary<string, LatexMacroDefinition>(StringComparer.Ordinal);
        foreach (LatexMacroDefinition definition in definitions.Where(static definition => definition.IsSafe)) {
            if (string.Equals(definition.Command.Name, "providecommand", StringComparison.Ordinal) && map.ContainsKey(definition.Name)) continue;
            map[definition.Name] = definition;
        }
        var diagnostics = new List<LatexMacroExpansionDiagnostic>();
        string output = ExpandCore(value, map, diagnostics, new HashSet<string>(StringComparer.Ordinal), 0, maximumDepth, maximumOutputLength);
        return new LatexMacroExpansionResult(output, diagnostics);
    }

    private static string ExpandCore(
        string value,
        IReadOnlyDictionary<string, LatexMacroDefinition> definitions,
        List<LatexMacroExpansionDiagnostic> diagnostics,
        HashSet<string> active,
        int depth,
        int maximumDepth,
        int maximumOutputLength) {
        if (depth > maximumDepth) throw new InvalidDataException("Simple macro expansion exceeds maximumDepth.");
        var output = new StringBuilder(value.Length);
        for (int index = 0; index < value.Length;) {
            if (value[index] != '\\' || index + 1 >= value.Length || !IsLetter(value[index + 1])) {
                output.Append(value[index++]);
                EnforceLength(output, maximumOutputLength);
                continue;
            }
            int nameStart = index + 1;
            int nameEnd = nameStart;
            while (nameEnd < value.Length && IsLetter(value[nameEnd])) nameEnd++;
            string name = value.Substring(nameStart, nameEnd - nameStart);
            if (!definitions.TryGetValue(name, out LatexMacroDefinition? definition)) {
                output.Append(value, index, nameEnd - index);
                index = nameEnd;
                continue;
            }
            if (!active.Add(name)) {
                diagnostics.Add(new LatexMacroExpansionDiagnostic("LATEXMAC002", LatexDiagnosticSeverity.Error,
                    "Cyclic simple macro invocation '" + name + "'.", index));
                output.Append(value, index, nameEnd - index);
                index = nameEnd;
                continue;
            }

            int cursor = nameEnd;
            var arguments = new List<string>();
            if (definition.DefaultValue != null) {
                SkipWhitespace(value, ref cursor);
                if (cursor < value.Length && value[cursor] == '[' && TryReadBalanced(value, ref cursor, '[', ']', out string optional)) {
                    arguments.Add(optional);
                } else {
                    arguments.Add(definition.DefaultValue);
                }
            }
            while (arguments.Count < definition.ParameterCount) {
                SkipWhitespace(value, ref cursor);
                if (cursor >= value.Length || value[cursor] != '{' || !TryReadBalanced(value, ref cursor, '{', '}', out string argument)) break;
                arguments.Add(argument);
            }
            if (arguments.Count != definition.ParameterCount) {
                diagnostics.Add(new LatexMacroExpansionDiagnostic("LATEXMAC001", LatexDiagnosticSeverity.Warning,
                    "Simple macro '" + name + "' did not receive the required arguments.", index));
                output.Append(value, index, nameEnd - index);
                active.Remove(name);
                index = nameEnd;
                continue;
            }
            string replacement = SubstituteParameters(definition.Body, arguments);
            output.Append(ExpandCore(replacement, definitions, diagnostics, active, depth + 1, maximumDepth, maximumOutputLength));
            active.Remove(name);
            index = cursor;
            EnforceLength(output, maximumOutputLength);
        }
        return output.ToString();
    }

    private static string SubstituteParameters(string body, IReadOnlyList<string> arguments) {
        var output = new StringBuilder(body.Length);
        for (int index = 0; index < body.Length; index++) {
            if (body[index] == '#' && index + 1 < body.Length && body[index + 1] >= '1' && body[index + 1] <= '9') {
                int parameter = body[index + 1] - '1';
                if (parameter < arguments.Count) output.Append(arguments[parameter]);
                index++;
            } else {
                output.Append(body[index]);
            }
        }
        return output.ToString();
    }

    private static bool TryReadBalanced(string value, ref int cursor, char open, char close, out string content) {
        content = string.Empty;
        if (cursor >= value.Length || value[cursor] != open) return false;
        int start = ++cursor;
        int depth = 1;
        while (cursor < value.Length) {
            if (value[cursor] == '\\') { cursor += Math.Min(2, value.Length - cursor); continue; }
            if (value[cursor] == open) depth++;
            else if (value[cursor] == close && --depth == 0) {
                content = value.Substring(start, cursor - start);
                cursor++;
                return true;
            }
            cursor++;
        }
        return false;
    }

    private static void SkipWhitespace(string value, ref int cursor) {
        while (cursor < value.Length && char.IsWhiteSpace(value[cursor])) cursor++;
    }

    private static void EnforceLength(StringBuilder output, int maximumOutputLength) {
        if (output.Length > maximumOutputLength) throw new InvalidDataException("Simple macro expansion exceeds maximumOutputLength.");
    }

    private static bool IsLetter(char value) =>
        (value >= 'a' && value <= 'z') || (value >= 'A' && value <= 'Z') || value == '@';
}
