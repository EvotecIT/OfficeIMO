using PdfExcel = OfficeIMO.Pdf.Excel;
using OfficeIMO.Pdf.Excel;
using OfficeIMO.Excel.OpenDocument;
using ExcelPdf = OfficeIMO.Excel.Pdf;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OpenDocument.Ods.Pdf;

/// <summary>Direct, loss-aware ODS to PDF conversion through the Excel semantic and PDF engines.</summary>
public static partial class OdsPdfConversionExtensions {
    /// <summary>Converts an ODS workbook to PDF and preserves diagnostics from both conversion stages.</summary>
    private static PdfCore.PdfDocumentConversionResult ConvertToPdfCore(this OdsDocument source, ExcelOpenDocumentConversionOptions? conversionOptions = null, ExcelPdf.ExcelPdfExportOptions? pdfOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (source == null) throw new ArgumentNullException(nameof(source));
        OdfConversionResult<OfficeIMO.Excel.ExcelDocument> conversion = source.ConvertToExcelWithEvidence(conversionOptions);
        using (conversion.Value) {
            PdfCore.PdfDocumentConversionResult result = ExcelPdf.ExcelPdfConverterExtensions.ConvertToPdfWithEvidence(conversion.Value, pdfOptions, cancellationToken);
            return result.WithSourceConversionReport(conversion.Report);
        }
    }

    /// <summary>Reconstructs an ODS workbook and preserves diagnostics from both table-conversion stages.</summary>
    private static PdfOdsConversionResult ConvertOdsCore(this PdfCore.PdfDocument source, PdfExcel.PdfExcelExportOptions? pdfOptions = null, ExcelOpenDocumentConversionOptions? openDocumentOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (source == null) throw new ArgumentNullException(nameof(source));
        PdfExcel.PdfExcelTableImportResult pdf = PdfExcel.PdfExcelTableConverterExtensions.ConvertToExcelWithEvidence(source, pdfOptions, cancellationToken);
        using (pdf.Value) {
            OdfConversionResult<OdsDocument> ods = pdf.Value.ConvertToOpenDocumentWithEvidence(openDocumentOptions);
            cancellationToken.ThrowIfCancellationRequested();
            return new PdfOdsConversionResult(
                ods.Value,
                new PdfOdsConversionReport(pdf.Report, ods.Report));
        }
    }

    /// <summary>Reconstructs an ODS workbook from a logical PDF model and preserves both stage reports.</summary>
    private static PdfOdsConversionResult ConvertOdsCore(this PdfCore.PdfDocumentReadResult source, PdfExcel.PdfExcelExportOptions? pdfOptions = null, ExcelOpenDocumentConversionOptions? openDocumentOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (source == null) throw new ArgumentNullException(nameof(source));
        PdfExcel.PdfExcelTableImportResult pdf = PdfExcel.PdfExcelTableConverterExtensions.ConvertToExcelWithEvidence(source, pdfOptions, cancellationToken);
        using (pdf.Value) {
            OdfConversionResult<OdsDocument> ods = pdf.Value.ConvertToOpenDocumentWithEvidence(openDocumentOptions);
            cancellationToken.ThrowIfCancellationRequested();
            return new PdfOdsConversionResult(
                ods.Value,
                new PdfOdsConversionReport(pdf.Report, ods.Report));
        }
    }

}
