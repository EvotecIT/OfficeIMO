using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

/// <summary>Immutable diagnostics from one semantic PDF-to-RTF conversion.</summary>
public sealed class PdfRtfConversionReport : OfficeConversionReport {
    internal PdfRtfConversionReport(PdfCore.PdfConversionReport report) {
        if (report == null) throw new ArgumentNullException(nameof(report));
        Warnings = Array.AsReadOnly(report.Warnings.ToArray());
    }

    /// <summary>Diagnostics captured while reconstructing editable RTF content.</summary>
    public IReadOnlyList<PdfCore.PdfConversionWarning> Warnings { get; }

    /// <summary>Category-preserving PDF reconstruction diagnostics.</summary>
    protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => FidelityDiagnostics;
    private IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> FidelityDiagnostics => Array.AsReadOnly(
        Warnings.Select(static warning => warning.ToFidelityDiagnostic()).ToArray());

    /// <summary>True when conversion reported an approximation, omission, or failure.</summary>
    public override bool HasLoss => Warnings.Any(static warning =>
        warning.LossKind != OfficeConversionLossKind.None);

    /// <summary>Throws when the conversion reported possible content loss.</summary>
    protected override OfficeConversionException CreateLossException() => new OfficeConversionException("PDF-to-RTF conversion reported possible content loss. First diagnostic: " + Warnings[0], this);
}

/// <summary>Editable RTF output and immutable diagnostics from one semantic PDF import.</summary>
public sealed class PdfRtfConversionResult : OfficeConversionResult<RtfDocument> {
    /// <summary>Gets the typed conversion report.</summary>
    public new PdfRtfConversionReport Report => (PdfRtfConversionReport)base.Report;

    internal PdfRtfConversionResult(RtfDocument value, PdfCore.PdfConversionReport report)
        : base(value, new PdfRtfConversionReport(report)) { }
}
