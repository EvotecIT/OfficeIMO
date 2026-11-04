using PdfExcel = OfficeIMO.Pdf.Excel;
using OfficeIMO.Excel.OpenDocument;
using ExcelPdf = OfficeIMO.Excel.Pdf;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OpenDocument.Ods.Pdf;

/// <summary>Controls OpenDocument spreadsheet to PDF conversion through the OpenDocument and PDF engines.</summary>
public sealed class OdsPdfExportOptions {
    /// <summary>OpenDocument import settings for the intermediate Office document.</summary>
    public ExcelOpenDocumentConversionOptions? ConversionOptions { get; set; }

    /// <summary>PDF export settings for the intermediate Office document.</summary>
    public ExcelPdf.ExcelPdfExportOptions? PdfOptions { get; set; }

    /// <summary>Throws <see cref="OfficeConversionException"/> when either stage reports loss.</summary>
    public bool FailOnLoss { get; set; }

    /// <summary>Creates a copy of these settings with independent PDF settings.</summary>
    public OdsPdfExportOptions Clone() => new OdsPdfExportOptions {
        ConversionOptions = ConversionOptions,
        PdfOptions = PdfOptions?.Clone(),
        FailOnLoss = FailOnLoss
    };
}

/// <summary>Controls PDF to OpenDocument spreadsheet reconstruction through the PDF and OpenDocument engines.</summary>
public sealed class PdfOdsExportOptions {
    /// <summary>PDF reconstruction settings for the intermediate Office document.</summary>
    public PdfExcel.PdfExcelExportOptions? PdfOptions { get; set; }

    /// <summary>OpenDocument export settings for the reconstructed document.</summary>
    public ExcelOpenDocumentConversionOptions? OpenDocumentOptions { get; set; }

    /// <summary>Throws <see cref="OfficeConversionException"/> when either stage reports loss.</summary>
    public bool FailOnLoss { get; set; }

    /// <summary>Creates a copy of these settings with independent PDF settings.</summary>
    public PdfOdsExportOptions Clone() => new PdfOdsExportOptions {
        PdfOptions = PdfOptions?.Clone(),
        OpenDocumentOptions = OpenDocumentOptions,
        FailOnLoss = FailOnLoss
    };
}
