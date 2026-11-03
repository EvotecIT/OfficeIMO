using System.Collections.Generic;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Word;

/// <summary>Immutable diagnostics from one semantic PDF-to-Word conversion.</summary>
public sealed class PdfWordConversionReport : OfficeConversionReport {
    internal PdfWordConversionReport(PdfCore.PdfConversionReport report) {
        ArgumentNullException.ThrowIfNull(report);
        Warnings = Array.AsReadOnly(report.Warnings.ToArray());
    }

    /// <summary>Diagnostics captured while reconstructing editable Word content.</summary>
    public IReadOnlyList<PdfCore.PdfConversionWarning> Warnings { get; }

    /// <summary>Category-preserving PDF reconstruction diagnostics.</summary>
    protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => FidelityDiagnostics;
    private IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> FidelityDiagnostics => Array.AsReadOnly(
        Warnings.Select(static warning => warning.ToFidelityDiagnostic()).ToArray());

    /// <summary>True when conversion reported an approximation, omission, or failure.</summary>
    public override bool HasLoss => Warnings.Any(static warning =>
        warning.LossKind != OfficeConversionLossKind.None);

}

/// <summary>Editable Word output and immutable diagnostics from one semantic PDF import.</summary>
public sealed class PdfWordConversionResult : OfficeConversionResult<WordDocument> {
    /// <summary>Gets the typed conversion report.</summary>
    public new PdfWordConversionReport Report => (PdfWordConversionReport)base.Report;

    internal PdfWordConversionResult(WordDocument value, PdfCore.PdfConversionReport report)
        : base(value, new PdfWordConversionReport(report)) { }
}
