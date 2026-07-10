namespace OfficeIMO.AsciiDoc;

/// <summary>Behavior when attribute substitution encounters an unset name.</summary>
public enum AsciiDocUndefinedAttributeBehavior {
    /// <summary>Keep the original <c>{name}</c> reference.</summary>
    Preserve = 0,
    /// <summary>Remove the reference.</summary>
    Drop,
    /// <summary>Keep the reference and report an error.</summary>
    Error
}

/// <summary>Bounded attribute-reference substitution options.</summary>
public sealed class AsciiDocAttributeSubstitutionOptions {
    /// <summary>Maximum recursive expansion passes.</summary>
    public int MaximumDepth { get; set; } = 32;

    /// <summary>Maximum produced character count.</summary>
    public int MaximumOutputLength { get; set; } = 64 * 1024 * 1024;

    /// <summary>Unset reference behavior.</summary>
    public AsciiDocUndefinedAttributeBehavior UndefinedAttributeBehavior { get; set; } = AsciiDocUndefinedAttributeBehavior.Preserve;
}

/// <summary>Diagnostic emitted during explicit AsciiDoc evaluation.</summary>
public sealed class AsciiDocEvaluationDiagnostic {
    internal AsciiDocEvaluationDiagnostic(string code, AsciiDocDiagnosticSeverity severity, string message, int offset) {
        Code = code;
        Severity = severity;
        Message = message;
        Offset = offset;
    }

    /// <summary>Stable diagnostic code.</summary>
    public string Code { get; }

    /// <summary>Diagnostic severity.</summary>
    public AsciiDocDiagnosticSeverity Severity { get; }

    /// <summary>Human-readable message.</summary>
    public string Message { get; }

    /// <summary>UTF-16 offset in the evaluated input.</summary>
    public int Offset { get; }
}

/// <summary>Result of bounded attribute substitution.</summary>
public sealed class AsciiDocAttributeSubstitutionResult {
    internal AsciiDocAttributeSubstitutionResult(string value, IReadOnlyList<AsciiDocEvaluationDiagnostic> diagnostics) {
        Value = value;
        Diagnostics = diagnostics;
    }

    /// <summary>Evaluated text.</summary>
    public string Value { get; }

    /// <summary>Cycles, missing values, and limit diagnostics.</summary>
    public IReadOnlyList<AsciiDocEvaluationDiagnostic> Diagnostics { get; }
}

/// <summary>Deterministic, bounded document attribute substitution.</summary>
public static class AsciiDocAttributeSubstitutor {
    /// <summary>Substitutes <c>{name}</c> references using explicit attributes.</summary>
    public static AsciiDocAttributeSubstitutionResult Substitute(
        string value,
        AsciiDocDocumentAttributes attributes,
        AsciiDocAttributeSubstitutionOptions? options = null) {
        if (value == null) throw new ArgumentNullException(nameof(value));
        if (attributes == null) throw new ArgumentNullException(nameof(attributes));
        options ??= new AsciiDocAttributeSubstitutionOptions();
        ValidateOptions(options);

        var diagnostics = new List<AsciiDocEvaluationDiagnostic>();
        var active = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        string output = Expand(value, attributes, options, diagnostics, active, 0, 0);
        return new AsciiDocAttributeSubstitutionResult(output, diagnostics);
    }

    private static string Expand(
        string value,
        AsciiDocDocumentAttributes attributes,
        AsciiDocAttributeSubstitutionOptions options,
        List<AsciiDocEvaluationDiagnostic> diagnostics,
        HashSet<string> active,
        int depth,
        int baseOffset) {
        if (depth > options.MaximumDepth) {
            diagnostics.Add(new AsciiDocEvaluationDiagnostic("ADOCEVAL003", AsciiDocDiagnosticSeverity.Error,
                "Attribute expansion exceeded MaximumDepth.", baseOffset));
            return value;
        }

        var output = new StringBuilder(value.Length);
        for (int index = 0; index < value.Length;) {
            if (value[index] == '\\' && index + 1 < value.Length && value[index + 1] == '{') {
                output.Append('{');
                index += 2;
                continue;
            }
            if (value[index] != '{') {
                output.Append(value[index++]);
                EnforceLength(output, options);
                continue;
            }
            int close = value.IndexOf('}', index + 1);
            if (close < 0) {
                output.Append(value, index, value.Length - index);
                break;
            }
            string name = value.Substring(index + 1, close - index - 1);
            if (!AsciiDocText.IsAttributeName(name)) {
                output.Append(value, index, close - index + 1);
                index = close + 1;
                continue;
            }
            if (!attributes.TryGetValue(name, out string replacement)) {
                HandleUndefined(output, value, index, close, name, options, diagnostics, baseOffset);
                index = close + 1;
                continue;
            }
            if (!active.Add(name)) {
                diagnostics.Add(new AsciiDocEvaluationDiagnostic("ADOCEVAL002", AsciiDocDiagnosticSeverity.Error,
                    "Cyclic document attribute reference '" + name + "'.", baseOffset + index));
                output.Append(value, index, close - index + 1);
                index = close + 1;
                continue;
            }
            output.Append(Expand(replacement, attributes, options, diagnostics, active, depth + 1, baseOffset + index));
            active.Remove(name);
            EnforceLength(output, options);
            index = close + 1;
        }
        EnforceLength(output, options);
        return output.ToString();
    }

    private static void HandleUndefined(
        StringBuilder output,
        string input,
        int start,
        int close,
        string name,
        AsciiDocAttributeSubstitutionOptions options,
        List<AsciiDocEvaluationDiagnostic> diagnostics,
        int baseOffset) {
        if (options.UndefinedAttributeBehavior != AsciiDocUndefinedAttributeBehavior.Drop) {
            output.Append(input, start, close - start + 1);
        }
        AsciiDocDiagnosticSeverity severity = options.UndefinedAttributeBehavior == AsciiDocUndefinedAttributeBehavior.Error
            ? AsciiDocDiagnosticSeverity.Error
            : AsciiDocDiagnosticSeverity.Warning;
        diagnostics.Add(new AsciiDocEvaluationDiagnostic("ADOCEVAL001", severity,
            "Document attribute '" + name + "' is not set.", baseOffset + start));
    }

    private static void EnforceLength(StringBuilder output, AsciiDocAttributeSubstitutionOptions options) {
        if (output.Length > options.MaximumOutputLength) throw new InvalidDataException("Attribute expansion exceeds MaximumOutputLength.");
    }

    private static void ValidateOptions(AsciiDocAttributeSubstitutionOptions options) {
        if (options.MaximumDepth < 1) throw new ArgumentOutOfRangeException(nameof(options), "MaximumDepth must be positive.");
        if (options.MaximumOutputLength < 1) throw new ArgumentOutOfRangeException(nameof(options), "MaximumOutputLength must be positive.");
    }
}
