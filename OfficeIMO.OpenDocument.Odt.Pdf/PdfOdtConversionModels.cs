using OfficeIMO.Word.Pdf;

namespace OfficeIMO.OpenDocument.Odt.Pdf;

/// <summary>Diagnostics from the PDF-to-Word and Word-to-ODT stages.</summary>
public sealed class PdfOdtConversionReport : IOfficeConversionReport {
    internal PdfOdtConversionReport(PdfWordConversionReport pdfReport, OdfConversionReport openDocumentReport) {
        PdfReport = pdfReport ?? throw new ArgumentNullException(nameof(pdfReport));
        OpenDocumentReport = openDocumentReport ?? throw new ArgumentNullException(nameof(openDocumentReport));
    }

    /// <summary>Diagnostics from semantic PDF-to-Word reconstruction.</summary>
    public PdfWordConversionReport PdfReport { get; }
    /// <summary>Feature mappings from Word to ODT.</summary>
    public OdfConversionReport OpenDocumentReport { get; }
    /// <summary>True when either stage reported possible loss.</summary>
    public bool HasLoss => PdfReport.HasLoss || OpenDocumentReport.HasLoss;

    /// <summary>Throws when either conversion stage reported possible loss.</summary>
    public void RequireNoLoss() {
        PdfReport.RequireNoLoss();
        OpenDocumentReport.RequireNoLoss();
    }
}

/// <summary>An ODT document with diagnostics from both conversion stages.</summary>
public sealed class PdfOdtConversionResult : OfficeConversionResult<OdtDocument, PdfOdtConversionReport> {
    internal PdfOdtConversionResult(OdtDocument value, PdfOdtConversionReport report) : base(value, report) { }
}
