using System.Collections.Generic;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

/// <summary>Immutable diagnostics from one semantic PDF-to-Word conversion.</summary>
public sealed class PdfWordConversionReport {
    internal PdfWordConversionReport(PdfCore.PdfConversionReport report) {
        if (report == null) throw new ArgumentNullException(nameof(report));
        Warnings = Array.AsReadOnly(report.Warnings.ToArray());
    }

    /// <summary>Diagnostics captured while reconstructing editable Word content.</summary>
    public IReadOnlyList<PdfCore.PdfConversionWarning> Warnings { get; }

    /// <summary>True when the conversion reported a warning or error severity diagnostic.</summary>
    public bool HasLoss => Warnings.Any(static warning =>
        warning.Severity != PdfCore.PdfConversionWarningSeverity.Information);

    /// <summary>Throws when the conversion reported possible content loss.</summary>
    public void RequireNoLoss() {
        if (HasLoss) {
            throw new InvalidOperationException("PDF-to-Word conversion reported possible content loss. First diagnostic: " + Warnings[0]);
        }
    }
}

/// <summary>Editable Word output and immutable diagnostics from one semantic PDF import.</summary>
public sealed class PdfWordConversionResult {
    internal PdfWordConversionResult(WordDocument value, PdfCore.PdfConversionReport report) {
        Value = value ?? throw new ArgumentNullException(nameof(value));
        Report = new PdfWordConversionReport(report);
    }

    /// <summary>The imported Word document.</summary>
    public WordDocument Value { get; }

    /// <summary>Snapshot of diagnostics reported by the import.</summary>
    public PdfWordConversionReport Report { get; }

    /// <summary>True when the conversion reported possible content loss.</summary>
    public bool HasLoss => Report.HasLoss;

    /// <summary>Returns the imported Word document.</summary>
    public WordDocument RequireValue() => Value;

    /// <summary>Returns the imported Word document only when no possible content loss was reported.</summary>
    public WordDocument RequireNoLoss() {
        Report.RequireNoLoss();
        return Value;
    }
}
