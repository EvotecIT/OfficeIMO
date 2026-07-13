namespace OfficeIMO.Latex.Markdown;

/// <summary>Conversion outcome.</summary>
public enum LatexMarkdownConversionOutcome {
    /// <summary>Equivalent typed target semantics.</summary>
    Converted = 0,
    /// <summary>Useful but less expressive target semantics.</summary>
    Simplified,
    /// <summary>Original source retained visibly.</summary>
    SourceFallback,
    /// <summary>Source omitted.</summary>
    Omitted
}

/// <summary>Loss-aware conversion diagnostic.</summary>
public sealed class LatexMarkdownConversionDiagnostic {
    internal LatexMarkdownConversionDiagnostic(
        string code,
        LatexMarkdownConversionOutcome outcome,
        string feature,
        string message,
        LatexSourceSpan? latexSpan,
        MarkdownSourceSpan? markdownSpan = null) {
        Code = code;
        Outcome = outcome;
        Feature = feature;
        Message = message;
        LatexSpan = latexSpan;
        MarkdownSpan = markdownSpan;
    }
    /// <summary>Stable code.</summary>
    public string Code { get; }
    /// <summary>Conversion outcome.</summary>
    public LatexMarkdownConversionOutcome Outcome { get; }
    /// <summary>Source feature.</summary>
    public string Feature { get; }
    /// <summary>Explanation.</summary>
    public string Message { get; }
    /// <summary>LaTeX source span when converting from LaTeX.</summary>
    public LatexSourceSpan? LatexSpan { get; }
    /// <summary>Markdown source span when converting from Markdown.</summary>
    public MarkdownSourceSpan? MarkdownSpan { get; }
}

/// <summary>LaTeX-to-Markdown options.</summary>
public sealed class LatexToMarkdownOptions {
    /// <summary>Preserves unsupported source in visible <c>latex</c> code blocks.</summary>
    public bool PreserveUnsupportedAsSource { get; set; } = true;
    /// <summary>Includes preamble title, author, date, and document class as YAML front matter.</summary>
    public bool IncludePreambleAsFrontMatter { get; set; } = true;
}

/// <summary>Markdown-to-LaTeX options.</summary>
public sealed class MarkdownToLatexOptions {
    /// <summary>Generated document class.</summary>
    public string DocumentClass { get; set; } = "article";
    /// <summary>Output line ending.</summary>
    public string LineEnding { get; set; } = "\n";
    /// <summary>Promotes the first H1 to <c>\title</c> and <c>\maketitle</c>.</summary>
    public bool FirstHeadingIsTitle { get; set; } = true;
    /// <summary>Preserves unsupported Markdown in a verbatim environment.</summary>
    public bool PreserveUnsupportedAsSource { get; set; } = true;
}

/// <summary>Markdown result from LaTeX.</summary>
public sealed class LatexToMarkdownResult {
    internal LatexToMarkdownResult(MarkdownDoc value, IReadOnlyList<LatexMarkdownConversionDiagnostic> diagnostics) {
        Value = value ?? throw new ArgumentNullException(nameof(value));
        Diagnostics = Array.AsReadOnly(diagnostics.ToArray());
    }
    /// <summary>Converted Markdown document.</summary>
    public MarkdownDoc Value { get; }
    /// <summary>Loss diagnostics.</summary>
    public IReadOnlyList<LatexMarkdownConversionDiagnostic> Diagnostics { get; }
    /// <summary>True when any feature was not exactly converted.</summary>
    public bool HasLoss => Diagnostics.Any(static diagnostic => diagnostic.Outcome != LatexMarkdownConversionOutcome.Converted);
}

/// <summary>Canonical LaTeX result from Markdown.</summary>
public sealed class MarkdownToLatexResult {
    internal MarkdownToLatexResult(string source, LatexDocument value, IReadOnlyList<LatexMarkdownConversionDiagnostic> diagnostics) {
        Source = source;
        Value = value ?? throw new ArgumentNullException(nameof(value));
        Diagnostics = Array.AsReadOnly(diagnostics.ToArray());
    }
    /// <summary>Generated source.</summary>
    public string Source { get; }
    /// <summary>Lossless parsed generated source.</summary>
    public LatexDocument Value { get; }
    /// <summary>Loss diagnostics.</summary>
    public IReadOnlyList<LatexMarkdownConversionDiagnostic> Diagnostics { get; }
    /// <summary>True when any feature was not exactly converted.</summary>
    public bool HasLoss => Diagnostics.Any(static diagnostic => diagnostic.Outcome != LatexMarkdownConversionOutcome.Converted);
}
