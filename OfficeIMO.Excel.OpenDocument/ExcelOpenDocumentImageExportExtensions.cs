using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.OpenDocument;
using System.Threading;

namespace OfficeIMO.Excel.OpenDocument;

/// <summary>Thin ODS image-export bridge over the Excel visual renderer.</summary>
public static class ExcelOpenDocumentImageExportExtensions {
    /// <summary>Converts an ODS workbook to Excel semantics and exports selected worksheets.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ExportImages(
        this OdsDocument source,
        OfficeImageExportFormat format,
        ExcelWorkbookImageExportOptions? imageOptions = null,
        ExcelOpenDocumentConversionOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) {
        var results = new List<OfficeImageExportResult>();
        source.ExportImages(
            format,
            results.Add,
            imageOptions,
            conversionOptions,
            cancellationToken);
        return results.AsReadOnly();
    }

    /// <summary>Streams selected ODS worksheet images without retaining earlier payloads.</summary>
    public static void ExportImages(
        this OdsDocument source,
        OfficeImageExportFormat format,
        OfficeImageExportConsumer consumer,
        ExcelWorkbookImageExportOptions? imageOptions = null,
        ExcelOpenDocumentConversionOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (consumer == null) throw new ArgumentNullException(nameof(consumer));
        ExcelWorkbookImageExportOptions effective =
            imageOptions?.CloneWorkbook() ?? new ExcelWorkbookImageExportOptions();
        OdfConversionResult<ExcelDocument> conversion =
            source.ToExcelDocumentResult(conversionOptions);
        using (conversion.Value) {
            conversion.Value.ExportImages(
                format,
                result => consumer(effective.EnsureAccepted(
                    OdfImageExportDiagnostics.Attach(result, conversion.Report))),
                effective,
                cancellationToken);
        }
    }
}
