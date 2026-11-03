using OfficeIMO.Drawing;

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
public sealed class LatexMarkdownExportOptions {
    /// <summary>Preserves unsupported source in visible <c>latex</c> code blocks.</summary>
    public bool PreserveUnsupportedAsSource { get; set; } = true;
    /// <summary>Includes preamble title, author, date, and document class as YAML front matter.</summary>
    public bool IncludePreambleAsFrontMatter { get; set; } = true;
}

/// <summary>Markdown-to-LaTeX options.</summary>
public sealed class MarkdownLatexExportOptions {
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
public sealed class LatexToMarkdownResult : OfficeConversionResult<MarkdownDocument> {
    /// <summary>Gets the typed conversion report.</summary>
    public new LatexToMarkdownReport Report => (LatexToMarkdownReport)base.Report;

    internal LatexToMarkdownResult(MarkdownDocument value, IReadOnlyList<LatexMarkdownConversionDiagnostic> diagnostics)
        : base(value, new LatexToMarkdownReport(diagnostics)) { }
}

/// <summary>Canonical LaTeX result from Markdown.</summary>
public sealed class MarkdownToLatexResult : OfficeConversionResult<LatexDocument> {
    /// <summary>Gets the typed conversion report.</summary>
    public new MarkdownToLatexReport Report => (MarkdownToLatexReport)base.Report;

    internal MarkdownToLatexResult(string source, LatexDocument value, IReadOnlyList<LatexMarkdownConversionDiagnostic> diagnostics)
        : base(value, new MarkdownToLatexReport(diagnostics)) {
        Source = source ?? throw new ArgumentNullException(nameof(source));
    }
    /// <summary>Generated source.</summary>
    public string Source { get; }
}

/// <summary>LaTeX-to-Markdown conversion diagnostics captured for one operation.</summary>
public sealed class LatexToMarkdownReport : OfficeConversionReport {
    internal LatexToMarkdownReport(IReadOnlyList<LatexMarkdownConversionDiagnostic> diagnostics) {
        NativeDiagnostics = Array.AsReadOnly((diagnostics ?? throw new ArgumentNullException(nameof(diagnostics))).ToArray());
        FidelityDiagnostics = CreateFidelityDiagnostics(NativeDiagnostics, "OfficeIMO.Latex.Markdown");
    }
    /// <summary>Loss, fallback, and omission diagnostics.</summary>
    public IReadOnlyList<LatexMarkdownConversionDiagnostic> NativeDiagnostics { get; }
    /// <summary>Category-preserving conversion diagnostics.</summary>
    protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => FidelityDiagnostics;
    private IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> FidelityDiagnostics { get; }
    /// <summary>True when at least one feature was not converted exactly.</summary>
    public override bool HasLoss => NativeDiagnostics.Any(static diagnostic => diagnostic.Outcome != LatexMarkdownConversionOutcome.Converted);
    /// <summary>Throws when the conversion reported a lossy mapping.</summary>
    protected override OfficeConversionException CreateLossException() => new OfficeConversionException("LaTeX-to-Markdown conversion reported one or more lossy mappings.", this);

    internal static IReadOnlyList<OfficeDiagnostic> CreateFidelityDiagnostics(
        IReadOnlyList<LatexMarkdownConversionDiagnostic> diagnostics,
        string source) => Array.AsReadOnly(diagnostics.Select(diagnostic =>
            new global::OfficeIMO.OfficeDiagnostic(
                diagnostic.Code,
                global::OfficeIMO.OfficeDiagnosticSeverity.Warning,
                diagnostic.Message,
                diagnostic.LatexSpan?.ToString() ?? diagnostic.MarkdownSpan?.ToString(),
                source,
                diagnostic.Outcome switch {
                    LatexMarkdownConversionOutcome.Converted => OfficeConversionLossKind.None,
                    LatexMarkdownConversionOutcome.Omitted => OfficeConversionLossKind.Omission,
                    _ => OfficeConversionLossKind.Approximation
                })).ToArray());
}

/// <summary>Markdown-to-LaTeX conversion diagnostics captured for one operation.</summary>
public sealed class MarkdownToLatexReport : OfficeConversionReport {
    internal MarkdownToLatexReport(IReadOnlyList<LatexMarkdownConversionDiagnostic> diagnostics) {
        NativeDiagnostics = Array.AsReadOnly((diagnostics ?? throw new ArgumentNullException(nameof(diagnostics))).ToArray());
        FidelityDiagnostics = LatexToMarkdownReport.CreateFidelityDiagnostics(NativeDiagnostics, "OfficeIMO.Latex.Markdown");
    }
    /// <summary>Loss, fallback, and omission diagnostics.</summary>
    public IReadOnlyList<LatexMarkdownConversionDiagnostic> NativeDiagnostics { get; }
    /// <summary>Category-preserving conversion diagnostics.</summary>
    protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => FidelityDiagnostics;
    private IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> FidelityDiagnostics { get; }
    /// <summary>True when at least one feature was not converted exactly.</summary>
    public override bool HasLoss => NativeDiagnostics.Any(static diagnostic => diagnostic.Outcome != LatexMarkdownConversionOutcome.Converted);
    /// <summary>Throws when the conversion reported a lossy mapping.</summary>
    protected override OfficeConversionException CreateLossException() => new OfficeConversionException("Markdown-to-LaTeX conversion reported one or more lossy mappings.", this);
}
