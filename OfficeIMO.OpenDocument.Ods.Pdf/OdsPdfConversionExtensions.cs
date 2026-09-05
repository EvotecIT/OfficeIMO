using OfficeIMO.Excel.OpenDocument;
using ExcelPdf = OfficeIMO.Excel.Pdf;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OpenDocument.Ods.Pdf;

/// <summary>Direct, loss-aware ODS to PDF conversion through the Excel semantic and PDF engines.</summary>
public static class OdsPdfConversionExtensions {
    /// <summary>Converts an ODS workbook to the first-party PDF document model.</summary>
    public static PdfCore.PdfDocument ToPdfDocument(this OdsDocument source, ExcelOpenDocumentConversionOptions? conversionOptions = null, ExcelPdf.ExcelToPdfOptions? pdfOptions = null, System.Threading.CancellationToken cancellationToken = default) =>
        source.ToPdfDocumentResult(conversionOptions, pdfOptions, cancellationToken).Value;

    /// <summary>Converts an ODS workbook to PDF and preserves diagnostics from both conversion stages.</summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this OdsDocument source, ExcelOpenDocumentConversionOptions? conversionOptions = null, ExcelPdf.ExcelToPdfOptions? pdfOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (source == null) throw new ArgumentNullException(nameof(source));
        OdfConversionResult<OfficeIMO.Excel.ExcelDocument> conversion = source.ToExcelDocumentResult(conversionOptions);
        using (conversion.Value) {
            PdfCore.PdfDocumentConversionResult result = ExcelPdf.ExcelPdfConverterExtensions.ToPdfDocumentResult(conversion.Value, pdfOptions, cancellationToken);
            return result.WithSourceConversionReport(conversion.Report);
        }
    }

    /// <summary>Converts an ODS workbook to PDF bytes.</summary>
    public static byte[] ToPdfBytes(this OdsDocument source, ExcelOpenDocumentConversionOptions? conversionOptions = null, ExcelPdf.ExcelToPdfOptions? pdfOptions = null, System.Threading.CancellationToken cancellationToken = default) =>
        source.ToPdfDocumentResult(conversionOptions, pdfOptions, cancellationToken).ToBytes(cancellationToken);

    /// <summary>Saves an ODS workbook as PDF.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this OdsDocument source, string path, ExcelOpenDocumentConversionOptions? conversionOptions = null, ExcelPdf.ExcelToPdfOptions? pdfOptions = null, System.Threading.CancellationToken cancellationToken = default) =>
        source.ToPdfDocumentResult(conversionOptions, pdfOptions, cancellationToken).Save(path, cancellationToken);

    /// <summary>Writes an ODS workbook as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this OdsDocument source, Stream stream, ExcelOpenDocumentConversionOptions? conversionOptions = null, ExcelPdf.ExcelToPdfOptions? pdfOptions = null, System.Threading.CancellationToken cancellationToken = default) =>
        source.ToPdfDocumentResult(conversionOptions, pdfOptions, cancellationToken).Save(stream, cancellationToken);

    /// <summary>Attempts to save an ODS workbook as PDF and returns structured failure evidence.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdfResult(this OdsDocument source, string path, ExcelOpenDocumentConversionOptions? conversionOptions = null, ExcelPdf.ExcelToPdfOptions? pdfOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return source.ToPdfDocumentResult(conversionOptions, pdfOptions, cancellationToken).SaveResult(path, cancellationToken); }
        catch (OperationCanceledException) { throw; }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(path, ex); }
    }

    /// <summary>Attempts to write an ODS workbook as PDF and returns structured failure evidence.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdfResult(this OdsDocument source, Stream stream, ExcelOpenDocumentConversionOptions? conversionOptions = null, ExcelPdf.ExcelToPdfOptions? pdfOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return source.ToPdfDocumentResult(conversionOptions, pdfOptions, cancellationToken).SaveResult(stream, cancellationToken); }
        catch (OperationCanceledException) { throw; }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex); }
    }

    /// <summary>Converts synchronously, then asynchronously saves an ODS workbook as PDF.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(this OdsDocument source, string path, ExcelOpenDocumentConversionOptions? conversionOptions = null, ExcelPdf.ExcelToPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return source.ToPdfDocumentResult(conversionOptions, pdfOptions, cancellationToken).SaveAsync(path, cancellationToken);
    }

    /// <summary>Converts synchronously, then asynchronously writes an ODS workbook as PDF.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(this OdsDocument source, Stream stream, ExcelOpenDocumentConversionOptions? conversionOptions = null, ExcelPdf.ExcelToPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return source.ToPdfDocumentResult(conversionOptions, pdfOptions, cancellationToken).SaveAsync(stream, cancellationToken);
    }

    /// <summary>Attempts to asynchronously save an ODS workbook as PDF and returns structured failure evidence.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfResultAsync(this OdsDocument source, string path, ExcelOpenDocumentConversionOptions? conversionOptions = null, ExcelPdf.ExcelToPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return await source.ToPdfDocumentResult(conversionOptions, pdfOptions, cancellationToken).SaveResultAsync(path, cancellationToken).ConfigureAwait(false); }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(path, ex); }
    }

    /// <summary>Attempts to asynchronously write an ODS workbook as PDF and returns structured failure evidence.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfResultAsync(this OdsDocument source, Stream stream, ExcelOpenDocumentConversionOptions? conversionOptions = null, ExcelPdf.ExcelToPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return await source.ToPdfDocumentResult(conversionOptions, pdfOptions, cancellationToken).SaveResultAsync(stream, cancellationToken).ConfigureAwait(false); }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex); }
    }

    /// <summary>Reconstructs an ODS workbook from an opened PDF.</summary>
    public static OdsDocument ToOdsDocument(this PdfCore.PdfDocument source, ExcelPdf.PdfTablesToExcelOptions? pdfOptions = null, ExcelOpenDocumentConversionOptions? openDocumentOptions = null, System.Threading.CancellationToken cancellationToken = default) =>
        source.ToOdsDocumentResult(pdfOptions, openDocumentOptions, cancellationToken).Value;

    /// <summary>Reconstructs an ODS workbook and preserves diagnostics from both table-conversion stages.</summary>
    public static PdfOdsConversionResult ToOdsDocumentResult(this PdfCore.PdfDocument source, ExcelPdf.PdfTablesToExcelOptions? pdfOptions = null, ExcelOpenDocumentConversionOptions? openDocumentOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (source == null) throw new ArgumentNullException(nameof(source));
        ExcelPdf.PdfExcelTableImportResult pdf = ExcelPdf.PdfExcelTableConverterExtensions.ImportTablesToExcelDocumentResult(source, pdfOptions, cancellationToken);
        using (pdf.Value) {
            OdfConversionResult<OdsDocument> ods = pdf.Value.ToOpenDocumentResult(openDocumentOptions);
            return new PdfOdsConversionResult(
                ods.Value,
                new PdfOdsConversionReport(pdf.Report, ods.Report));
        }
    }

    /// <summary>Reconstructs an ODS workbook from an already loaded logical PDF model.</summary>
    public static OdsDocument ToOdsDocument(this PdfCore.PdfDocumentReadResult source, ExcelPdf.PdfTablesToExcelOptions? pdfOptions = null, ExcelOpenDocumentConversionOptions? openDocumentOptions = null, System.Threading.CancellationToken cancellationToken = default) =>
        source.ToOdsDocumentResult(pdfOptions, openDocumentOptions, cancellationToken).Value;

    /// <summary>Reconstructs an ODS workbook from a logical PDF model and preserves both stage reports.</summary>
    public static PdfOdsConversionResult ToOdsDocumentResult(this PdfCore.PdfDocumentReadResult source, ExcelPdf.PdfTablesToExcelOptions? pdfOptions = null, ExcelOpenDocumentConversionOptions? openDocumentOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (source == null) throw new ArgumentNullException(nameof(source));
        ExcelPdf.PdfExcelTableImportResult pdf = ExcelPdf.PdfExcelTableConverterExtensions.ImportTablesToExcelDocumentResult(source, pdfOptions, cancellationToken);
        using (pdf.Value) {
            OdfConversionResult<OdsDocument> ods = pdf.Value.ToOpenDocumentResult(openDocumentOptions);
            return new PdfOdsConversionResult(
                ods.Value,
                new PdfOdsConversionReport(pdf.Report, ods.Report));
        }
    }

    /// <summary>Reconstructs and saves an ODS workbook from an opened PDF.</summary>
    public static OfficeOutputResult<PdfOdsConversionReport> SaveAsOds(this PdfCore.PdfDocument source, string path, ExcelPdf.PdfTablesToExcelOptions? pdfOptions = null, ExcelOpenDocumentConversionOptions? openDocumentOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdsConversionResult result = source.ToOdsDocumentResult(pdfOptions, openDocumentOptions, cancellationToken);
        result.Value.Save(path);
        return OfficeOutputResult<PdfOdsConversionReport>.FromSuccess(path, result.Report);
    }

    /// <summary>Reconstructs and writes an ODS workbook from an opened PDF.</summary>
    public static OfficeOutputResult<PdfOdsConversionReport> SaveAsOds(this PdfCore.PdfDocument source, Stream stream, ExcelPdf.PdfTablesToExcelOptions? pdfOptions = null, ExcelOpenDocumentConversionOptions? openDocumentOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdsConversionResult result = source.ToOdsDocumentResult(pdfOptions, openDocumentOptions, cancellationToken);
        result.Value.Save(stream);
        return OfficeOutputResult<PdfOdsConversionReport>.FromSuccess(null, result.Report);
    }

    /// <summary>Reconstructs and saves an ODS workbook from a logical PDF model.</summary>
    public static OfficeOutputResult<PdfOdsConversionReport> SaveAsOds(this PdfCore.PdfDocumentReadResult source, string path, ExcelPdf.PdfTablesToExcelOptions? pdfOptions = null, ExcelOpenDocumentConversionOptions? openDocumentOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdsConversionResult result = source.ToOdsDocumentResult(pdfOptions, openDocumentOptions, cancellationToken);
        result.Value.Save(path);
        return OfficeOutputResult<PdfOdsConversionReport>.FromSuccess(path, result.Report);
    }

    /// <summary>Reconstructs and writes an ODS workbook from a logical PDF model.</summary>
    public static OfficeOutputResult<PdfOdsConversionReport> SaveAsOds(this PdfCore.PdfDocumentReadResult source, Stream stream, ExcelPdf.PdfTablesToExcelOptions? pdfOptions = null, ExcelOpenDocumentConversionOptions? openDocumentOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdsConversionResult result = source.ToOdsDocumentResult(pdfOptions, openDocumentOptions, cancellationToken);
        result.Value.Save(stream);
        return OfficeOutputResult<PdfOdsConversionReport>.FromSuccess(null, result.Report);
    }

    /// <summary>Reconstructs synchronously, then asynchronously saves an ODS workbook from an opened PDF.</summary>
    public static async Task<OfficeOutputResult<PdfOdsConversionReport>> SaveAsOdsAsync(this PdfCore.PdfDocument source, string path, ExcelPdf.PdfTablesToExcelOptions? pdfOptions = null, ExcelOpenDocumentConversionOptions? openDocumentOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdsConversionResult result = source.ToOdsDocumentResult(pdfOptions, openDocumentOptions, cancellationToken);
        await result.Value.SaveAsync(path, cancellationToken: cancellationToken).ConfigureAwait(false);
        return OfficeOutputResult<PdfOdsConversionReport>.FromSuccess(path, result.Report);
    }

    /// <summary>Reconstructs synchronously, then asynchronously writes an ODS workbook from an opened PDF.</summary>
    public static async Task<OfficeOutputResult<PdfOdsConversionReport>> SaveAsOdsAsync(this PdfCore.PdfDocument source, Stream stream, ExcelPdf.PdfTablesToExcelOptions? pdfOptions = null, ExcelOpenDocumentConversionOptions? openDocumentOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdsConversionResult result = source.ToOdsDocumentResult(pdfOptions, openDocumentOptions, cancellationToken);
        await result.Value.SaveAsync(stream, cancellationToken: cancellationToken).ConfigureAwait(false);
        return OfficeOutputResult<PdfOdsConversionReport>.FromSuccess(null, result.Report);
    }

    /// <summary>Reconstructs synchronously, then asynchronously saves an ODS workbook from a logical PDF model.</summary>
    public static async Task<OfficeOutputResult<PdfOdsConversionReport>> SaveAsOdsAsync(this PdfCore.PdfDocumentReadResult source, string path, ExcelPdf.PdfTablesToExcelOptions? pdfOptions = null, ExcelOpenDocumentConversionOptions? openDocumentOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdsConversionResult result = source.ToOdsDocumentResult(pdfOptions, openDocumentOptions, cancellationToken);
        await result.Value.SaveAsync(path, cancellationToken: cancellationToken).ConfigureAwait(false);
        return OfficeOutputResult<PdfOdsConversionReport>.FromSuccess(path, result.Report);
    }

    /// <summary>Reconstructs synchronously, then asynchronously writes an ODS workbook from a logical PDF model.</summary>
    public static async Task<OfficeOutputResult<PdfOdsConversionReport>> SaveAsOdsAsync(this PdfCore.PdfDocumentReadResult source, Stream stream, ExcelPdf.PdfTablesToExcelOptions? pdfOptions = null, ExcelOpenDocumentConversionOptions? openDocumentOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdsConversionResult result = source.ToOdsDocumentResult(pdfOptions, openDocumentOptions, cancellationToken);
        await result.Value.SaveAsync(stream, cancellationToken: cancellationToken).ConfigureAwait(false);
        return OfficeOutputResult<PdfOdsConversionReport>.FromSuccess(null, result.Report);
    }
}
