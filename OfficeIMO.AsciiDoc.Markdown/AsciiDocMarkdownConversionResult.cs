using OfficeIMO.Drawing;

namespace OfficeIMO.AsciiDoc.Markdown;

/// <summary>Markdown document plus explicit conversion diagnostics.</summary>
public sealed class AsciiDocToMarkdownResult : OfficeConversionResult<MarkdownDoc, AsciiDocToMarkdownReport> {
    internal AsciiDocToMarkdownResult(MarkdownDoc value, IReadOnlyList<AsciiDocMarkdownConversionDiagnostic> diagnostics)
        : base(value, new AsciiDocToMarkdownReport(diagnostics)) { }
}

/// <summary>AsciiDoc-to-Markdown conversion diagnostics captured for one operation.</summary>
public sealed class AsciiDocToMarkdownReport : IOfficeConversionReport {
    internal AsciiDocToMarkdownReport(IReadOnlyList<AsciiDocMarkdownConversionDiagnostic> diagnostics) {
        Diagnostics = Array.AsReadOnly((diagnostics ?? throw new ArgumentNullException(nameof(diagnostics))).ToArray());
        FidelityDiagnostics = Array.AsReadOnly(Diagnostics.Select(static diagnostic =>
            new OfficeConversionFidelityDiagnostic(
                diagnostic.Code,
                diagnostic.Message,
                diagnostic.Outcome switch {
                    AsciiDocMarkdownConversionOutcome.Converted => OfficeConversionLossKind.None,
                    AsciiDocMarkdownConversionOutcome.Omitted => OfficeConversionLossKind.Omission,
                    _ => OfficeConversionLossKind.Approximation
                },
                "OfficeIMO.AsciiDoc.Markdown",
                diagnostic.SourceSpan.ToString())).ToArray());
    }

    /// <summary>Loss, fallback, and omission diagnostics.</summary>
    public IReadOnlyList<AsciiDocMarkdownConversionDiagnostic> Diagnostics { get; }

    /// <summary>Category-preserving conversion diagnostics.</summary>
    public IReadOnlyList<OfficeConversionFidelityDiagnostic> FidelityDiagnostics { get; }

    /// <summary>True when at least one feature was not converted exactly.</summary>
    public bool HasLoss => Diagnostics.Any(static diagnostic => diagnostic.Outcome != AsciiDocMarkdownConversionOutcome.Converted);

    /// <summary>Throws when the conversion reported a lossy mapping.</summary>
    public void RequireNoLoss() {
        if (HasLoss) throw new InvalidOperationException("AsciiDoc-to-Markdown conversion reported one or more lossy mappings.");
    }
}
