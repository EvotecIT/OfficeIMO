using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.OpenDocument;

namespace OfficeIMO.Excel.OpenDocument;

public static partial class ExcelOpenDocumentImageExportExtensions {
    /// <summary>Renders selected source surfaces using the encoding in the image options.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ToImages(this OdsDocument document,
        ExcelImageExportOptions? imageOptions = null, ExcelOpenDocumentConversionOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) =>
        CreateImageOperation(document, imageOptions?.Clone() ?? new ExcelImageExportOptions(), conversionOptions).Export(cancellationToken);

    /// <summary>Renders selected source surfaces asynchronously under one shared operation deadline and budget.</summary>
    public static Task<IReadOnlyList<OfficeImageExportResult>> ToImagesAsync(this OdsDocument document,
        ExcelImageExportOptions? imageOptions = null, ExcelOpenDocumentConversionOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) {
        var resolved = imageOptions?.Clone() ?? new ExcelImageExportOptions();
        return CreateImageOperation(document, resolved, conversionOptions).ExportAsync(cancellationToken);
    }

    /// <summary>Saves selected source surfaces through the shared atomic file and fidelity policies.</summary>
    public static OfficeConversionReport SaveAsImages(this OdsDocument document, string folderPath,
        ExcelImageExportOptions? imageOptions = null, ExcelOpenDocumentConversionOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        var resolved = imageOptions?.Clone() ?? new ExcelImageExportOptions();
        var saved = CreateImageOperation(document, resolved, conversionOptions).SaveFiles(folderPath, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("ODS", resolved.Format, saved.Report.NativeDiagnostics,
            saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    /// <summary>Saves selected source surfaces asynchronously without retaining earlier encoded payloads.</summary>
    public static async Task<OfficeConversionReport> SaveAsImagesAsync(this OdsDocument document, string folderPath,
        ExcelImageExportOptions? imageOptions = null, ExcelOpenDocumentConversionOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        var resolved = imageOptions?.Clone() ?? new ExcelImageExportOptions();
        var saved = await CreateImageOperation(document, resolved, conversionOptions)
            .SaveFilesAsync(folderPath, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("ODS", resolved.Format, saved.Report.NativeDiagnostics,
            saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    private static OfficeImageExportBatchOperation<ExcelImageExportOptions> CreateImageOperation(
        OdsDocument document, ExcelImageExportOptions options, ExcelOpenDocumentConversionOptions? conversionOptions) {
        ArgumentNullException.ThrowIfNull(document);
        ExcelOpenDocumentConversionOptions? conversion = conversionOptions?.Clone();
        return new OfficeImageExportBatchOperation<ExcelImageExportOptions>(options, options.Format, options.FileConflictPolicy,
            (format, effective) => document.ToImages(format, effective, conversion),
            exportEach: (format, effective, consumer, token) => document.ToImages(format, consumer, effective, conversion, token));
    }
}
