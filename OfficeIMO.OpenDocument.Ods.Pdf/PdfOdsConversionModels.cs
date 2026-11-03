using OfficeIMO.Pdf.Excel;
using System.Collections.Generic;
using OfficeIMO.Excel.Pdf;

namespace OfficeIMO.OpenDocument.Ods.Pdf;

/// <summary>Diagnostics from the PDF-to-Excel and Excel-to-ODS stages.</summary>
public sealed class PdfOdsConversionReport : OfficeConversionReport {
    internal PdfOdsConversionReport(PdfExcelTableImportReport pdfReport, OdfConversionReport openDocumentReport) {
        PdfReport = pdfReport ?? throw new ArgumentNullException(nameof(pdfReport));
        OpenDocumentReport = openDocumentReport ?? throw new ArgumentNullException(nameof(openDocumentReport));
    }

    /// <summary>Detected-table import evidence and omitted PDF page scope.</summary>
    public PdfExcelTableImportReport PdfReport { get; }
    /// <summary>Feature mappings from Excel to ODS.</summary>
    public OdfConversionReport OpenDocumentReport { get; }
    /// <summary>Category-preserving diagnostics from both stages.</summary>
    protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => FidelityDiagnostics;
    private IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> FidelityDiagnostics =>
        OfficeConversionFidelityDiagnostics.Flatten(new OfficeConversionReport[] { PdfReport, OpenDocumentReport });
    /// <summary>True when either stage truncated, approximated, skipped, or omitted source content.</summary>
    public override bool HasLoss => PdfReport.HasLoss || PdfReport.HasOmittedPageContent || OpenDocumentReport.HasLoss;

}

/// <summary>An ODS document with diagnostics from both conversion stages.</summary>
public sealed class PdfOdsConversionResult : OfficeConversionResult<OdsDocument> {
    /// <summary>Gets the typed conversion report.</summary>
    public new PdfOdsConversionReport Report => (PdfOdsConversionReport)base.Report;

    internal PdfOdsConversionResult(OdsDocument value, PdfOdsConversionReport report) : base(value, report) { }
}
