using OfficeIMO.PowerPoint.Pdf;

namespace OfficeIMO.OpenDocument.Odp.Pdf;

/// <summary>Diagnostics from the PDF-to-PowerPoint and PowerPoint-to-ODP stages.</summary>
public sealed class PdfOdpConversionReport : IOfficeConversionReport {
    internal PdfOdpConversionReport(PdfPowerPointConversionReport pdfReport, OdfConversionReport openDocumentReport) {
        PdfReport = pdfReport ?? throw new ArgumentNullException(nameof(pdfReport));
        OpenDocumentReport = openDocumentReport ?? throw new ArgumentNullException(nameof(openDocumentReport));
    }

    /// <summary>Visual-page or editable-table PDF import evidence.</summary>
    public PdfPowerPointConversionReport PdfReport { get; }
    /// <summary>Feature mappings from PowerPoint to ODP.</summary>
    public OdfConversionReport OpenDocumentReport { get; }
    /// <summary>True when either stage truncated, approximated, skipped, or omitted source content.</summary>
    public bool HasLoss => PdfReport.HasLoss || PdfReport.HasOmittedPageContent || OpenDocumentReport.HasLoss;

    /// <summary>Throws when either stage reported possible loss or omitted PDF page content.</summary>
    public void RequireNoLoss() {
        PdfReport.RequireNoLoss();
        if (PdfReport.HasOmittedPageContent) {
            throw new InvalidOperationException("PDF-to-ODP conversion omitted page content outside detected tables.");
        }
        OpenDocumentReport.RequireNoLoss();
    }
}

/// <summary>An ODP presentation with diagnostics from both conversion stages.</summary>
public sealed class PdfOdpConversionResult : OfficeConversionResult<OdpPresentation, PdfOdpConversionReport> {
    internal PdfOdpConversionResult(OdpPresentation value, PdfOdpConversionReport report) : base(value, report) { }
}
