using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Markdown;

namespace OfficeIMO.Word.Markdown;

/// <summary>One structured diagnostic from a Word/Markdown conversion.</summary>
public sealed class WordMarkdownConversionDiagnostic {
    /// <summary>Creates a diagnostic.</summary>
    public WordMarkdownConversionDiagnostic(string code, string message, OfficeConversionLossKind lossKind) {
        Code = string.IsNullOrWhiteSpace(code) ? throw new ArgumentException("A diagnostic code is required.", nameof(code)) : code;
        Message = message ?? throw new ArgumentNullException(nameof(message));
        LossKind = lossKind;
    }

    /// <summary>Stable diagnostic code.</summary>
    public string Code { get; }

    /// <summary>Human-readable diagnostic message.</summary>
    public string Message { get; }

    /// <summary>Fidelity impact of the diagnostic.</summary>
    public OfficeConversionLossKind LossKind { get; }
}

/// <summary>Immutable fidelity report from one Word/Markdown conversion.</summary>
public sealed class WordMarkdownConversionReport : IOfficeConversionReport {
    /// <summary>Creates a report from conversion diagnostics.</summary>
    public WordMarkdownConversionReport(IEnumerable<WordMarkdownConversionDiagnostic>? diagnostics = null) {
        Diagnostics = Array.AsReadOnly((diagnostics ?? Array.Empty<WordMarkdownConversionDiagnostic>()).ToArray());
    }

    /// <summary>Structured diagnostics in emission order.</summary>
    public IReadOnlyList<WordMarkdownConversionDiagnostic> Diagnostics { get; }

    /// <summary>Whether conversion completed without a failure diagnostic.</summary>
    public bool Succeeded => !Diagnostics.Any(static diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Failure);

    /// <summary>Whether any source content was approximated, omitted, or failed.</summary>
    public bool HasLoss => Diagnostics.Any(static diagnostic => diagnostic.LossKind != OfficeConversionLossKind.None);

    /// <summary>Throws when the report contains fidelity loss.</summary>
    public void RequireNoLoss() {
        if (HasLoss) throw new WordMarkdownConversionException(this);
    }
}

/// <summary>Exception thrown when a Word/Markdown conversion is required to be lossless.</summary>
public sealed class WordMarkdownConversionException : InvalidOperationException {
    /// <summary>Creates an exception for a lossy report.</summary>
    public WordMarkdownConversionException(WordMarkdownConversionReport report)
        : base("The Word/Markdown conversion did not preserve all source content.") {
        Report = report ?? throw new ArgumentNullException(nameof(report));
    }

    /// <summary>Report that caused the exception.</summary>
    public WordMarkdownConversionReport Report { get; }
}

/// <summary>Shared value-and-report contract for Word/Markdown conversions.</summary>
public abstract class WordMarkdownConversionResult<T> : OfficeConversionResult<T, WordMarkdownConversionReport> where T : class {
    /// <summary>Creates a conversion result.</summary>
    protected WordMarkdownConversionResult(T value, WordMarkdownConversionReport report) : base(value, report) { }

    /// <summary>Whether conversion completed without a failure diagnostic.</summary>
    public override bool Succeeded => Report.Succeeded;

    /// <summary>Returns the value when conversion succeeded.</summary>
    public override T RequireValue() {
        if (!Succeeded) throw new WordMarkdownConversionException(Report);
        return base.RequireValue();
    }

    /// <summary>Returns the value only when conversion was lossless.</summary>
    public override T RequireNoLoss() {
        Report.RequireNoLoss();
        return base.RequireValue();
    }
}

/// <summary>Typed Markdown document plus the Word-to-Markdown fidelity report.</summary>
public sealed class WordToMarkdownResult : WordMarkdownConversionResult<MarkdownDoc> {
    internal WordToMarkdownResult(MarkdownDoc value, WordMarkdownConversionReport report) : base(value, report) { }
}

/// <summary>Native Word document plus the Markdown-to-Word fidelity report.</summary>
public sealed class MarkdownToWordResult : WordMarkdownConversionResult<WordDocument> {
    internal MarkdownToWordResult(WordDocument value, WordMarkdownConversionReport report) : base(value, report) { }
}
