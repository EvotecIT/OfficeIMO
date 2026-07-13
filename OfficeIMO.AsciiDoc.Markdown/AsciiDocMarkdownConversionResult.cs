namespace OfficeIMO.AsciiDoc.Markdown;

/// <summary>Markdown document plus explicit conversion diagnostics.</summary>
public sealed class AsciiDocToMarkdownResult {
    internal AsciiDocToMarkdownResult(MarkdownDoc value, IReadOnlyList<AsciiDocMarkdownConversionDiagnostic> diagnostics) {
        Value = value ?? throw new ArgumentNullException(nameof(value));
        Report = new AsciiDocToMarkdownReport(diagnostics);
    }

    /// <summary>Converted Markdown semantic document.</summary>
    public MarkdownDoc Value { get; }

    /// <summary>Snapshot of conversion diagnostics and loss state.</summary>
    public AsciiDocToMarkdownReport Report { get; }

    /// <summary>True when at least one feature was simplified, source-fallbacked, or omitted.</summary>
    public bool HasLoss => Report.HasLoss;

    /// <summary>Returns the converted document.</summary>
    public MarkdownDoc RequireValue() => Value;

    /// <summary>Returns the converted document only when no lossy mapping was reported.</summary>
    public MarkdownDoc RequireNoLoss() {
        Report.RequireNoLoss();
        return Value;
    }
}

/// <summary>AsciiDoc-to-Markdown conversion diagnostics captured for one operation.</summary>
public sealed class AsciiDocToMarkdownReport {
    internal AsciiDocToMarkdownReport(IReadOnlyList<AsciiDocMarkdownConversionDiagnostic> diagnostics) {
        Diagnostics = Array.AsReadOnly((diagnostics ?? throw new ArgumentNullException(nameof(diagnostics))).ToArray());
    }

    /// <summary>Loss, fallback, and omission diagnostics.</summary>
    public IReadOnlyList<AsciiDocMarkdownConversionDiagnostic> Diagnostics { get; }

    /// <summary>True when at least one feature was not converted exactly.</summary>
    public bool HasLoss => Diagnostics.Any(static diagnostic => diagnostic.Outcome != AsciiDocMarkdownConversionOutcome.Converted);

    /// <summary>Throws when the conversion reported a lossy mapping.</summary>
    public void RequireNoLoss() {
        if (HasLoss) throw new InvalidOperationException("AsciiDoc-to-Markdown conversion reported one or more lossy mappings.");
    }
}
