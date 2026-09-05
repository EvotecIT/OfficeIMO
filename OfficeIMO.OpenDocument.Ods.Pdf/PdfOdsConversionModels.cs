using OfficeIMO.Excel.Pdf;

namespace OfficeIMO.OpenDocument.Ods.Pdf;

/// <summary>Diagnostics from the PDF-to-Excel and Excel-to-ODS stages.</summary>
public sealed class PdfOdsConversionReport : IOfficeConversionReport {
    internal PdfOdsConversionReport(PdfExcelTableImportReport pdfReport, OdfConversionReport openDocumentReport) {
        PdfReport = pdfReport ?? throw new ArgumentNullException(nameof(pdfReport));
        OpenDocumentReport = openDocumentReport ?? throw new ArgumentNullException(nameof(openDocumentReport));
    }

    /// <summary>Detected-table import evidence and omitted PDF page scope.</summary>
    public PdfExcelTableImportReport PdfReport { get; }
    /// <summary>Feature mappings from Excel to ODS.</summary>
    public OdfConversionReport OpenDocumentReport { get; }
    /// <summary>True when either stage truncated, approximated, skipped, or omitted source content.</summary>
    public bool HasLoss => PdfReport.HasLoss || PdfReport.HasOmittedPageContent || OpenDocumentReport.HasLoss;

    /// <summary>Throws when either conversion stage reported possible loss or omitted PDF page content.</summary>
    public void RequireNoLoss() {
        PdfReport.RequireNoLoss();
        if (PdfReport.HasOmittedPageContent) {
            throw new InvalidOperationException("PDF-to-ODS conversion omitted page content outside detected tables.");
        }
        OpenDocumentReport.RequireNoLoss();
    }
}

/// <summary>An ODS document with diagnostics from both conversion stages.</summary>
public sealed class PdfOdsConversionResult : OfficeConversionResult<OdsDocument, PdfOdsConversionReport> {
    internal PdfOdsConversionResult(OdsDocument value, PdfOdsConversionReport report) : base(value, report) { }
}
