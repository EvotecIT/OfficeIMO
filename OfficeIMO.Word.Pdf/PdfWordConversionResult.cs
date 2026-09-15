using System.Collections.Generic;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

/// <summary>Immutable diagnostics from one semantic PDF-to-Word conversion.</summary>
public sealed class PdfWordConversionReport : IOfficeConversionReport {
    internal PdfWordConversionReport(PdfCore.PdfConversionReport report) {
        if (report == null) throw new ArgumentNullException(nameof(report));
        Warnings = Array.AsReadOnly(report.Warnings.ToArray());
    }

    /// <summary>Diagnostics captured while reconstructing editable Word content.</summary>
    public IReadOnlyList<PdfCore.PdfConversionWarning> Warnings { get; }

    /// <summary>True when conversion reported an approximation, omission, or failure.</summary>
    public bool HasLoss => Warnings.Any(static warning =>
        warning.LossKind != OfficeConversionLossKind.None);

    /// <summary>Throws when the conversion reported possible content loss.</summary>
    public void RequireNoLoss() {
        PdfCore.PdfConversionWarning? firstLoss = Warnings.FirstOrDefault(static warning =>
            warning.LossKind != OfficeConversionLossKind.None);
        if (firstLoss != null) {
            throw new InvalidOperationException("PDF-to-Word conversion reported possible content loss. First diagnostic: " + firstLoss);
        }
    }
}

/// <summary>Editable Word output and immutable diagnostics from one semantic PDF import.</summary>
public sealed class PdfWordConversionResult : OfficeConversionResult<WordDocument, PdfWordConversionReport> {
    internal PdfWordConversionResult(WordDocument value, PdfCore.PdfConversionReport report)
        : base(value, new PdfWordConversionReport(report)) { }
}
