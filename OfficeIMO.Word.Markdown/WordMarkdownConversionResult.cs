using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Markdown;

namespace OfficeIMO.Word.Markdown;

/// <summary>Fidelity impact represented by a Word/Markdown conversion diagnostic.</summary>
public enum WordMarkdownConversionLossKind {
    /// <summary>The diagnostic does not indicate fidelity loss.</summary>
    None = 0,
    /// <summary>The source was represented using an approximation or fallback.</summary>
    Approximation = 1,
    /// <summary>Source content was omitted.</summary>
    Omission = 2,
    /// <summary>The requested conversion could not be completed.</summary>
    Failure = 3
}

/// <summary>One structured diagnostic from a Word/Markdown conversion.</summary>
public sealed class WordMarkdownConversionDiagnostic {
    /// <summary>Creates a diagnostic.</summary>
    public WordMarkdownConversionDiagnostic(string code, string message, WordMarkdownConversionLossKind lossKind) {
        Code = string.IsNullOrWhiteSpace(code) ? throw new ArgumentException("A diagnostic code is required.", nameof(code)) : code;
        Message = message ?? throw new ArgumentNullException(nameof(message));
        LossKind = lossKind;
    }

    /// <summary>Stable diagnostic code.</summary>
    public string Code { get; }

    /// <summary>Human-readable diagnostic message.</summary>
    public string Message { get; }

    /// <summary>Fidelity impact of the diagnostic.</summary>
    public WordMarkdownConversionLossKind LossKind { get; }
}

/// <summary>Immutable fidelity report from one Word/Markdown conversion.</summary>
public sealed class WordMarkdownConversionReport {
    /// <summary>Creates a report from conversion diagnostics.</summary>
    public WordMarkdownConversionReport(IEnumerable<WordMarkdownConversionDiagnostic>? diagnostics = null) {
        Diagnostics = Array.AsReadOnly((diagnostics ?? Array.Empty<WordMarkdownConversionDiagnostic>()).ToArray());
    }

    /// <summary>Structured diagnostics in emission order.</summary>
    public IReadOnlyList<WordMarkdownConversionDiagnostic> Diagnostics { get; }

    /// <summary>Whether conversion completed without a failure diagnostic.</summary>
    public bool Succeeded => !Diagnostics.Any(static diagnostic => diagnostic.LossKind == WordMarkdownConversionLossKind.Failure);

    /// <summary>Whether any source content was approximated, omitted, or failed.</summary>
    public bool HasLoss => Diagnostics.Any(static diagnostic => diagnostic.LossKind != WordMarkdownConversionLossKind.None);

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
public abstract class WordMarkdownConversionResult<T> {
    /// <summary>Creates a conversion result.</summary>
    protected WordMarkdownConversionResult(T value, WordMarkdownConversionReport report) {
        Value = value ?? throw new ArgumentNullException(nameof(value));
        Report = report ?? throw new ArgumentNullException(nameof(report));
    }

    /// <summary>Native target value.</summary>
    public T Value { get; }

    /// <summary>Conversion fidelity report.</summary>
    public WordMarkdownConversionReport Report { get; }

    /// <summary>Whether conversion completed without a failure diagnostic.</summary>
    public bool Succeeded => Report.Succeeded;

    /// <summary>Whether conversion introduced fidelity loss.</summary>
    public bool HasLoss => Report.HasLoss;

    /// <summary>Returns the value when conversion succeeded.</summary>
    public T RequireValue() {
        if (!Succeeded) throw new WordMarkdownConversionException(Report);
        return Value;
    }

    /// <summary>Returns the value only when conversion was lossless.</summary>
    public T RequireNoLoss() {
        Report.RequireNoLoss();
        return Value;
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
