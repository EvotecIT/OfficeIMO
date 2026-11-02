using System.Collections.Generic;
using OfficeIMO.Word.Pdf;

namespace OfficeIMO.OpenDocument.Odt.Pdf;

/// <summary>Diagnostics from the PDF-to-Word and Word-to-ODT stages.</summary>
public sealed class PdfOdtConversionReport : OfficeConversionReport {
    internal PdfOdtConversionReport(PdfWordConversionReport pdfReport, OdfConversionReport openDocumentReport) {
        PdfReport = pdfReport ?? throw new ArgumentNullException(nameof(pdfReport));
        OpenDocumentReport = openDocumentReport ?? throw new ArgumentNullException(nameof(openDocumentReport));
    }

    /// <summary>Diagnostics from semantic PDF-to-Word reconstruction.</summary>
    public PdfWordConversionReport PdfReport { get; }
    /// <summary>Feature mappings from Word to ODT.</summary>
    public OdfConversionReport OpenDocumentReport { get; }
    /// <summary>Category-preserving diagnostics from both stages.</summary>
    protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => FidelityDiagnostics;
    private IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> FidelityDiagnostics =>
        OfficeConversionFidelityDiagnostics.Flatten(new OfficeConversionReport[] { PdfReport, OpenDocumentReport });
    /// <summary>True when either stage reported possible loss.</summary>
    public override bool HasLoss => PdfReport.HasLoss || OpenDocumentReport.HasLoss;

}

/// <summary>An ODT document with diagnostics from both conversion stages.</summary>
public sealed class PdfOdtConversionResult : OfficeConversionResult<OdtDocument> {
    /// <summary>Gets the typed conversion report.</summary>
    public new PdfOdtConversionReport Report => (PdfOdtConversionReport)base.Report;

    internal PdfOdtConversionResult(OdtDocument value, PdfOdtConversionReport report) : base(value, report) { }
}
