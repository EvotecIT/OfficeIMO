using OfficeIMO.Excel.OpenDocument;
using ExcelPdf = OfficeIMO.Excel.Pdf;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OpenDocument.Ods.Pdf;

/// <summary>Direct, loss-aware ODS to PDF conversion through the Excel semantic and PDF engines.</summary>
public static class OdsPdfConversionExtensions {
    /// <summary>Converts an ODS workbook to the first-party PDF document model.</summary>
    public static PdfCore.PdfDocument ToPdfDocument(this OdsDocument source, ExcelOpenDocumentConversionOptions? conversionOptions = null, ExcelPdf.ExcelPdfSaveOptions? pdfOptions = null) =>
        source.ToPdfDocumentResult(conversionOptions, pdfOptions).Value;

    /// <summary>Converts an ODS workbook to PDF and preserves diagnostics from both conversion stages.</summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this OdsDocument source, ExcelOpenDocumentConversionOptions? conversionOptions = null, ExcelPdf.ExcelPdfSaveOptions? pdfOptions = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        OdfConversionResult<OfficeIMO.Excel.ExcelDocument> conversion = source.ToExcelDocumentResult(conversionOptions);
        using (conversion.Value) {
            PdfCore.PdfDocumentConversionResult result = ExcelPdf.ExcelPdfConverterExtensions.ToPdfDocumentResult(conversion.Value, pdfOptions);
            return result.WithSourceConversionReport(conversion.Report);
        }
    }

    /// <summary>Converts an ODS workbook to PDF bytes.</summary>
    public static byte[] ToPdf(this OdsDocument source, ExcelOpenDocumentConversionOptions? conversionOptions = null, ExcelPdf.ExcelPdfSaveOptions? pdfOptions = null) =>
        source.ToPdfDocumentResult(conversionOptions, pdfOptions).ToBytes();

    /// <summary>Saves an ODS workbook as PDF.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this OdsDocument source, string path, ExcelOpenDocumentConversionOptions? conversionOptions = null, ExcelPdf.ExcelPdfSaveOptions? pdfOptions = null) =>
        source.ToPdfDocumentResult(conversionOptions, pdfOptions).Save(path);

    /// <summary>Writes an ODS workbook as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this OdsDocument source, Stream stream, ExcelOpenDocumentConversionOptions? conversionOptions = null, ExcelPdf.ExcelPdfSaveOptions? pdfOptions = null) =>
        source.ToPdfDocumentResult(conversionOptions, pdfOptions).Save(stream);

    /// <summary>Attempts to save an ODS workbook as PDF and returns structured failure evidence.</summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this OdsDocument source, string path, ExcelOpenDocumentConversionOptions? conversionOptions = null, ExcelPdf.ExcelPdfSaveOptions? pdfOptions = null) {
        try { return source.ToPdfDocumentResult(conversionOptions, pdfOptions).TrySave(path); }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(path, ex); }
    }

    /// <summary>Attempts to write an ODS workbook as PDF and returns structured failure evidence.</summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this OdsDocument source, Stream stream, ExcelOpenDocumentConversionOptions? conversionOptions = null, ExcelPdf.ExcelPdfSaveOptions? pdfOptions = null) {
        try { return source.ToPdfDocumentResult(conversionOptions, pdfOptions).TrySave(stream); }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex); }
    }

    /// <summary>Converts synchronously, then asynchronously saves an ODS workbook as PDF.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(this OdsDocument source, string path, ExcelOpenDocumentConversionOptions? conversionOptions = null, ExcelPdf.ExcelPdfSaveOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return source.ToPdfDocumentResult(conversionOptions, pdfOptions).SaveAsync(path, cancellationToken);
    }

    /// <summary>Converts synchronously, then asynchronously writes an ODS workbook as PDF.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(this OdsDocument source, Stream stream, ExcelOpenDocumentConversionOptions? conversionOptions = null, ExcelPdf.ExcelPdfSaveOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return source.ToPdfDocumentResult(conversionOptions, pdfOptions).SaveAsync(stream, cancellationToken);
    }

    /// <summary>Attempts to asynchronously save an ODS workbook as PDF and returns structured failure evidence.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(this OdsDocument source, string path, ExcelOpenDocumentConversionOptions? conversionOptions = null, ExcelPdf.ExcelPdfSaveOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return await source.ToPdfDocumentResult(conversionOptions, pdfOptions).TrySaveAsync(path, cancellationToken).ConfigureAwait(false); }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(path, ex); }
    }

    /// <summary>Attempts to asynchronously write an ODS workbook as PDF and returns structured failure evidence.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(this OdsDocument source, Stream stream, ExcelOpenDocumentConversionOptions? conversionOptions = null, ExcelPdf.ExcelPdfSaveOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return await source.ToPdfDocumentResult(conversionOptions, pdfOptions).TrySaveAsync(stream, cancellationToken).ConfigureAwait(false); }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex); }
    }

    /// <summary>Reconstructs an ODS workbook from an opened PDF.</summary>
    public static OdsDocument ToOdsDocument(this PdfCore.PdfDocument source, ExcelPdf.PdfExcelTableImportOptions? pdfOptions = null, ExcelOpenDocumentConversionOptions? openDocumentOptions = null) =>
        source.ToOdsDocumentResult(pdfOptions, openDocumentOptions).Value;

    /// <summary>Reconstructs an ODS workbook and preserves diagnostics from both table-conversion stages.</summary>
    public static PdfOdsConversionResult ToOdsDocumentResult(this PdfCore.PdfDocument source, ExcelPdf.PdfExcelTableImportOptions? pdfOptions = null, ExcelOpenDocumentConversionOptions? openDocumentOptions = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        ExcelPdf.PdfExcelTableImportResult pdf = ExcelPdf.PdfExcelTableConverterExtensions.ImportTablesToExcelDocumentResult(source, pdfOptions);
        using (pdf.Value) {
            OdfConversionResult<OdsDocument> ods = pdf.Value.ToOpenDocumentResult(openDocumentOptions);
            return new PdfOdsConversionResult(
                ods.Value,
                new PdfOdsConversionReport(pdf.Report, ods.Report));
        }
    }

    /// <summary>Reconstructs an ODS workbook from an already loaded logical PDF model.</summary>
    public static OdsDocument ToOdsDocument(this PdfCore.PdfDocumentReadResult source, ExcelPdf.PdfExcelTableImportOptions? pdfOptions = null, ExcelOpenDocumentConversionOptions? openDocumentOptions = null) =>
        source.ToOdsDocumentResult(pdfOptions, openDocumentOptions).Value;

    /// <summary>Reconstructs an ODS workbook from a logical PDF model and preserves both stage reports.</summary>
    public static PdfOdsConversionResult ToOdsDocumentResult(this PdfCore.PdfDocumentReadResult source, ExcelPdf.PdfExcelTableImportOptions? pdfOptions = null, ExcelOpenDocumentConversionOptions? openDocumentOptions = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        ExcelPdf.PdfExcelTableImportResult pdf = ExcelPdf.PdfExcelTableConverterExtensions.ImportTablesToExcelDocumentResult(source, pdfOptions);
        using (pdf.Value) {
            OdfConversionResult<OdsDocument> ods = pdf.Value.ToOpenDocumentResult(openDocumentOptions);
            return new PdfOdsConversionResult(
                ods.Value,
                new PdfOdsConversionReport(pdf.Report, ods.Report));
        }
    }

    /// <summary>Reconstructs and saves an ODS workbook from an opened PDF.</summary>
    public static PdfOdsConversionReport SaveAsOds(this PdfCore.PdfDocument source, string path, ExcelPdf.PdfExcelTableImportOptions? pdfOptions = null, ExcelOpenDocumentConversionOptions? openDocumentOptions = null) {
        PdfOdsConversionResult result = source.ToOdsDocumentResult(pdfOptions, openDocumentOptions);
        result.Value.Save(path);
        return result.Report;
    }

    /// <summary>Reconstructs and writes an ODS workbook from an opened PDF.</summary>
    public static PdfOdsConversionReport SaveAsOds(this PdfCore.PdfDocument source, Stream stream, ExcelPdf.PdfExcelTableImportOptions? pdfOptions = null, ExcelOpenDocumentConversionOptions? openDocumentOptions = null) {
        PdfOdsConversionResult result = source.ToOdsDocumentResult(pdfOptions, openDocumentOptions);
        result.Value.Save(stream);
        return result.Report;
    }

    /// <summary>Reconstructs and saves an ODS workbook from a logical PDF model.</summary>
    public static PdfOdsConversionReport SaveAsOds(this PdfCore.PdfDocumentReadResult source, string path, ExcelPdf.PdfExcelTableImportOptions? pdfOptions = null, ExcelOpenDocumentConversionOptions? openDocumentOptions = null) {
        PdfOdsConversionResult result = source.ToOdsDocumentResult(pdfOptions, openDocumentOptions);
        result.Value.Save(path);
        return result.Report;
    }

    /// <summary>Reconstructs and writes an ODS workbook from a logical PDF model.</summary>
    public static PdfOdsConversionReport SaveAsOds(this PdfCore.PdfDocumentReadResult source, Stream stream, ExcelPdf.PdfExcelTableImportOptions? pdfOptions = null, ExcelOpenDocumentConversionOptions? openDocumentOptions = null) {
        PdfOdsConversionResult result = source.ToOdsDocumentResult(pdfOptions, openDocumentOptions);
        result.Value.Save(stream);
        return result.Report;
    }

    /// <summary>Reconstructs synchronously, then asynchronously saves an ODS workbook from an opened PDF.</summary>
    public static async Task<PdfOdsConversionReport> SaveAsOdsAsync(this PdfCore.PdfDocument source, string path, ExcelPdf.PdfExcelTableImportOptions? pdfOptions = null, ExcelOpenDocumentConversionOptions? openDocumentOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdsConversionResult result = source.ToOdsDocumentResult(pdfOptions, openDocumentOptions);
        await result.Value.SaveAsync(path, cancellationToken: cancellationToken).ConfigureAwait(false);
        return result.Report;
    }

    /// <summary>Reconstructs synchronously, then asynchronously writes an ODS workbook from an opened PDF.</summary>
    public static async Task<PdfOdsConversionReport> SaveAsOdsAsync(this PdfCore.PdfDocument source, Stream stream, ExcelPdf.PdfExcelTableImportOptions? pdfOptions = null, ExcelOpenDocumentConversionOptions? openDocumentOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdsConversionResult result = source.ToOdsDocumentResult(pdfOptions, openDocumentOptions);
        await result.Value.SaveAsync(stream, cancellationToken: cancellationToken).ConfigureAwait(false);
        return result.Report;
    }

    /// <summary>Reconstructs synchronously, then asynchronously saves an ODS workbook from a logical PDF model.</summary>
    public static async Task<PdfOdsConversionReport> SaveAsOdsAsync(this PdfCore.PdfDocumentReadResult source, string path, ExcelPdf.PdfExcelTableImportOptions? pdfOptions = null, ExcelOpenDocumentConversionOptions? openDocumentOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdsConversionResult result = source.ToOdsDocumentResult(pdfOptions, openDocumentOptions);
        await result.Value.SaveAsync(path, cancellationToken: cancellationToken).ConfigureAwait(false);
        return result.Report;
    }

    /// <summary>Reconstructs synchronously, then asynchronously writes an ODS workbook from a logical PDF model.</summary>
    public static async Task<PdfOdsConversionReport> SaveAsOdsAsync(this PdfCore.PdfDocumentReadResult source, Stream stream, ExcelPdf.PdfExcelTableImportOptions? pdfOptions = null, ExcelOpenDocumentConversionOptions? openDocumentOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdsConversionResult result = source.ToOdsDocumentResult(pdfOptions, openDocumentOptions);
        await result.Value.SaveAsync(stream, cancellationToken: cancellationToken).ConfigureAwait(false);
        return result.Report;
    }
}
