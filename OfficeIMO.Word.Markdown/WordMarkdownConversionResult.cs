using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Markdown;

namespace OfficeIMO.Word.Markdown;

/// <summary>One structured diagnostic from a Word/Markdown conversion.</summary>
internal sealed class WordMarkdownConversionDiagnostic {
    /// <summary>Creates a diagnostic.</summary>
    public WordMarkdownConversionDiagnostic(string code, string message, OfficeConversionLossKind lossKind,
        OfficeDiagnosticSeverity severity = OfficeDiagnosticSeverity.Warning, string? location = null, string? format = null) {
        Code = string.IsNullOrWhiteSpace(code) ? throw new ArgumentException("A diagnostic code is required.", nameof(code)) : code;
        Message = message ?? throw new ArgumentNullException(nameof(message));
        LossKind = lossKind;
        Severity = severity; Location = location; Format = format;
    }

    /// <summary>Finding severity, independent of fidelity loss.</summary>
    public OfficeDiagnosticSeverity Severity { get; }
    /// <summary>Source part, block or element location.</summary>
    public string? Location { get; }
    /// <summary>Format or stage that emitted the finding.</summary>
    public string? Format { get; }

    /// <summary>Stable diagnostic code.</summary>
    public string Code { get; }

    /// <summary>Human-readable diagnostic message.</summary>
    public string Message { get; }

    /// <summary>Fidelity impact of the diagnostic.</summary>
    public OfficeConversionLossKind LossKind { get; }
}

/// <summary>Immutable fidelity report from one Word/Markdown conversion.</summary>
internal sealed class WordMarkdownConversionReport : OfficeConversionReport {
    /// <summary>Creates a report from conversion diagnostics.</summary>
    public WordMarkdownConversionReport(IEnumerable<WordMarkdownConversionDiagnostic>? diagnostics = null) {
        NativeDiagnostics = Array.AsReadOnly((diagnostics ?? Array.Empty<WordMarkdownConversionDiagnostic>()).ToArray());
        FidelityDiagnostics = Array.AsReadOnly(NativeDiagnostics.Select(static diagnostic =>
            new global::OfficeIMO.OfficeDiagnostic(
                diagnostic.Code,
                diagnostic.LossKind == global::OfficeIMO.OfficeConversionLossKind.None ? global::OfficeIMO.OfficeDiagnosticSeverity.Information : global::OfficeIMO.OfficeDiagnosticSeverity.Warning,
                diagnostic.Message,
                null,
                "OfficeIMO.Word.Markdown",
                diagnostic.LossKind)).ToArray());
    }

    /// <summary>Structured diagnostics in emission order.</summary>
    public IReadOnlyList<WordMarkdownConversionDiagnostic> NativeDiagnostics { get; }

    /// <summary>Category-preserving conversion diagnostics.</summary>
    protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => FidelityDiagnostics;
    private IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> FidelityDiagnostics { get; }

    /// <summary>Whether conversion completed without a failure diagnostic.</summary>
    public bool Succeeded => !NativeDiagnostics.Any(static diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Failure);

    /// <summary>Whether any source content was approximated, omitted, or failed.</summary>
    public override bool HasLoss => NativeDiagnostics.Any(static diagnostic => diagnostic.LossKind != OfficeConversionLossKind.None);

    /// <summary>Throws when the report contains fidelity loss.</summary>
    protected override OfficeConversionException CreateLossException() => new WordMarkdownConversionException(this);
}

/// <summary>Exception thrown when a Word/Markdown conversion is required to be lossless.</summary>
internal sealed class WordMarkdownConversionException : OfficeConversionException {
    /// <summary>Creates an exception for a lossy report.</summary>
    public WordMarkdownConversionException(WordMarkdownConversionReport report)
        : base("The Word/Markdown conversion did not preserve all source content.", report) {
        Report = report;
    }

    /// <summary>Report that caused the exception.</summary>
    public new WordMarkdownConversionReport Report { get; }
}

/// <summary>Shared value-and-report contract for Word/Markdown conversions.</summary>
internal abstract class WordMarkdownConversionResult<T> : OfficeConversionResult<T> where T : class {
    /// <summary>Gets the typed conversion report.</summary>
    public new WordMarkdownConversionReport Report => (WordMarkdownConversionReport)base.Report;

    /// <summary>Creates a conversion result.</summary>
    protected WordMarkdownConversionResult(T value, WordMarkdownConversionReport report) : base(value, report) { }

    /// <summary>Whether conversion completed without a failure diagnostic.</summary>
    public bool Succeeded => Report.Succeeded;

    /// <summary>Returns the value when conversion succeeded.</summary>
    public T RequireValue() {
        if (!Succeeded) throw new WordMarkdownConversionException(Report);
        return Value;
    }

}

/// <summary>Typed Markdown document plus the Word-to-Markdown fidelity report.</summary>
internal sealed class WordToMarkdownResult : WordMarkdownConversionResult<MarkdownDocument> {
    internal WordToMarkdownResult(MarkdownDocument value, WordMarkdownConversionReport report) : base(value, report) { }
}

/// <summary>Native Word document plus the Markdown-to-Word fidelity report.</summary>
internal sealed class MarkdownToWordResult : WordMarkdownConversionResult<WordDocument> {
    internal MarkdownToWordResult(WordDocument value, WordMarkdownConversionReport report) : base(value, report) { }
}
