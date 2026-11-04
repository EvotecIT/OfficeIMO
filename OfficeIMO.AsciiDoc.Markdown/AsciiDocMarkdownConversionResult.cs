using OfficeIMO.Drawing;

namespace OfficeIMO.AsciiDoc.Markdown;

/// <summary>Markdown document plus explicit conversion diagnostics.</summary>
public sealed class AsciiDocToMarkdownResult : OfficeConversionResult<MarkdownDocument> {
    /// <summary>Gets the typed conversion report.</summary>
    public new AsciiDocToMarkdownReport Report => (AsciiDocToMarkdownReport)base.Report;

    internal AsciiDocToMarkdownResult(MarkdownDocument value, IReadOnlyList<AsciiDocMarkdownConversionDiagnostic> diagnostics)
        : base(value, new AsciiDocToMarkdownReport(diagnostics)) { }
}

/// <summary>AsciiDoc-to-Markdown conversion diagnostics captured for one operation.</summary>
public sealed class AsciiDocToMarkdownReport : OfficeConversionReport {
    internal AsciiDocToMarkdownReport(IReadOnlyList<AsciiDocMarkdownConversionDiagnostic> diagnostics) {
        NativeDiagnostics = Array.AsReadOnly((diagnostics ?? throw new ArgumentNullException(nameof(diagnostics))).ToArray());
        FidelityDiagnostics = Array.AsReadOnly(NativeDiagnostics.Select(static diagnostic =>
            new global::OfficeIMO.OfficeDiagnostic(
                diagnostic.Code,
                global::OfficeIMO.OfficeDiagnosticSeverity.Warning,
                diagnostic.Message,
                diagnostic.SourceSpan.ToString(),
                "OfficeIMO.AsciiDoc.Markdown",
                diagnostic.Outcome switch {
                    AsciiDocMarkdownConversionOutcome.Converted => OfficeConversionLossKind.None,
                    AsciiDocMarkdownConversionOutcome.Omitted => OfficeConversionLossKind.Omission,
                    _ => OfficeConversionLossKind.Approximation
                })).ToArray());
    }

    /// <summary>Loss, fallback, and omission diagnostics.</summary>
    public IReadOnlyList<AsciiDocMarkdownConversionDiagnostic> NativeDiagnostics { get; }

    /// <summary>Category-preserving conversion diagnostics.</summary>
    protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => FidelityDiagnostics;
    private IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> FidelityDiagnostics { get; }

    /// <summary>True when at least one feature was not converted exactly.</summary>
    public override bool HasLoss => NativeDiagnostics.Any(static diagnostic => diagnostic.Outcome != AsciiDocMarkdownConversionOutcome.Converted);

    /// <summary>Throws when the conversion reported a lossy mapping.</summary>
    protected override OfficeConversionException CreateLossException() => new OfficeConversionException("AsciiDoc-to-Markdown conversion reported one or more lossy mappings.", this);
}
