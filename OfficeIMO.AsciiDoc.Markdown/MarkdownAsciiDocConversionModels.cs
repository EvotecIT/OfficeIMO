namespace OfficeIMO.AsciiDoc.Markdown;

/// <summary>Options for canonical Markdown-to-AsciiDoc conversion.</summary>
public sealed class MarkdownToAsciiDocOptions {
    /// <summary>Output line ending. Defaults to LF.</summary>
    public string LineEnding { get; set; } = "\n";

    /// <summary>Treats the first Markdown H1 as the AsciiDoc document title. Defaults to true.</summary>
    public bool FirstLevelOneHeadingIsDocumentTitle { get; set; } = true;

    /// <summary>Preserves unsupported Markdown source inside a visible AsciiDoc listing block.</summary>
    public bool PreserveUnsupportedAsSource { get; set; } = true;
}

/// <summary>Diagnostic emitted while converting Markdown semantics to AsciiDoc.</summary>
public sealed class MarkdownAsciiDocConversionDiagnostic {
    internal MarkdownAsciiDocConversionDiagnostic(
        string code,
        AsciiDocMarkdownDiagnosticSeverity severity,
        AsciiDocMarkdownConversionOutcome outcome,
        string feature,
        string message,
        MarkdownSourceSpan? sourceSpan) {
        Code = code;
        Severity = severity;
        Outcome = outcome;
        Feature = feature;
        Message = message;
        SourceSpan = sourceSpan;
    }

    /// <summary>Stable diagnostic code.</summary>
    public string Code { get; }

    /// <summary>Severity.</summary>
    public AsciiDocMarkdownDiagnosticSeverity Severity { get; }

    /// <summary>Conversion outcome.</summary>
    public AsciiDocMarkdownConversionOutcome Outcome { get; }

    /// <summary>Source feature.</summary>
    public string Feature { get; }

    /// <summary>Human-readable explanation.</summary>
    public string Message { get; }

    /// <summary>Markdown source span when the input was parsed.</summary>
    public MarkdownSourceSpan? SourceSpan { get; }
}

/// <summary>Canonical AsciiDoc source and its lossless parsed document.</summary>
public sealed class MarkdownToAsciiDocResult {
    internal MarkdownToAsciiDocResult(
        string source,
        AsciiDocDocument value,
        IReadOnlyList<MarkdownAsciiDocConversionDiagnostic> diagnostics) {
        Source = source;
        Value = value ?? throw new ArgumentNullException(nameof(value));
        Diagnostics = Array.AsReadOnly(diagnostics.ToArray());
    }

    /// <summary>Generated canonical AsciiDoc.</summary>
    public string Source { get; }

    /// <summary>Lossless parsed view of <see cref="Source"/>.</summary>
    public AsciiDocDocument Value { get; }

    /// <summary>Fallback and simplification diagnostics.</summary>
    public IReadOnlyList<MarkdownAsciiDocConversionDiagnostic> Diagnostics { get; }

    /// <summary>True when any source feature was simplified, fallbacked, or omitted.</summary>
    public bool HasLoss => Diagnostics.Any(static diagnostic => diagnostic.Outcome != AsciiDocMarkdownConversionOutcome.Converted);
}
