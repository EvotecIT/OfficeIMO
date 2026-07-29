using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

/// <summary>Immutable diagnostics from one semantic PDF-to-RTF conversion.</summary>
public sealed class PdfRtfConversionReport {
    internal PdfRtfConversionReport(PdfCore.PdfConversionReport report) {
        if (report == null) throw new ArgumentNullException(nameof(report));
        Warnings = Array.AsReadOnly(report.Warnings.ToArray());
    }

    /// <summary>Diagnostics captured while reconstructing editable RTF content.</summary>
    public IReadOnlyList<PdfCore.PdfConversionWarning> Warnings { get; }

    /// <summary>True when the conversion reported a warning or error severity diagnostic.</summary>
    public bool HasLoss => Warnings.Any(static warning =>
        warning.Severity != PdfCore.PdfConversionWarningSeverity.Information);

    /// <summary>Throws when the conversion reported possible content loss.</summary>
    public void RequireNoLoss() {
        if (HasLoss) {
            throw new InvalidOperationException("PDF-to-RTF conversion reported possible content loss. First diagnostic: " + Warnings[0]);
        }
    }
}

/// <summary>Editable RTF output and immutable diagnostics from one semantic PDF import.</summary>
public sealed class PdfRtfConversionResult {
    internal PdfRtfConversionResult(RtfDocument value, PdfCore.PdfConversionReport report) {
        Value = value ?? throw new ArgumentNullException(nameof(value));
        Report = new PdfRtfConversionReport(report);
    }

    /// <summary>The imported RTF document.</summary>
    public RtfDocument Value { get; }

    /// <summary>Snapshot of diagnostics reported by the import.</summary>
    public PdfRtfConversionReport Report { get; }

    /// <summary>True when the conversion reported possible content loss.</summary>
    public bool HasLoss => Report.HasLoss;

    /// <summary>Returns the imported RTF document.</summary>
    public RtfDocument RequireValue() => Value;

    /// <summary>Returns the imported RTF document only when no possible content loss was reported.</summary>
    public RtfDocument RequireNoLoss() {
        Report.RequireNoLoss();
        return Value;
    }
}
