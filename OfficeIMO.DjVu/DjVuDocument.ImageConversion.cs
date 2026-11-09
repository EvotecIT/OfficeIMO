using System.Threading.Tasks;
using System.Diagnostics;
using OfficeIMO.Drawing;

namespace OfficeIMO.DjVu;

/// <summary>Renders DjVu surfaces through the existing native projection and shared image policies.</summary>
public static class DjVuDocumentImageConversionExtensions {
    /// <summary>Renders selected source pages in caller order, retaining source and encoding diagnostics.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ToImages(this DjVuDocument document,
        OfficeImageExportOptions? options = null, DjVuRenderOptions? renderOptions = null, CancellationToken cancellationToken = default) =>
        CreateOperation(document, options, renderOptions).Export(cancellationToken);

    /// <summary>Streams selected images without retaining earlier encoded payloads.</summary>
    public static void ToImages(this DjVuDocument document, OfficeImageExportConsumer consumer,
        OfficeImageExportOptions? options = null, DjVuRenderOptions? renderOptions = null, CancellationToken cancellationToken = default) =>
        CreateOperation(document, options, renderOptions).ExportEach(consumer, cancellationToken);

    /// <summary>Renders selected source pages asynchronously with shared cancellation and deadlines.</summary>
    public static Task<IReadOnlyList<OfficeImageExportResult>> ToImagesAsync(this DjVuDocument document,
        OfficeImageExportOptions? options = null, DjVuRenderOptions? renderOptions = null, CancellationToken cancellationToken = default) =>
        CreateOperation(document, options, renderOptions).ExportAsync(cancellationToken);

    /// <summary>Saves selected pages with atomic files, bounded output and fidelity acceptance.</summary>
    public static OfficeConversionReport SaveAsImages(this DjVuDocument document, string folderPath,
        OfficeImageExportOptions? options = null, DjVuRenderOptions? renderOptions = null, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        OfficeImageExportFormat format = options?.Format ?? OfficeImageExportFormat.Png;
        var operation = CreateOperation(document, options, renderOptions);
        var saved = operation.SaveFiles(folderPath, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("DjVu", format,
            saved.Report.NativeDiagnostics, saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    /// <summary>Saves selected pages asynchronously without retaining completed image payloads.</summary>
    public static async Task<OfficeConversionReport> SaveAsImagesAsync(this DjVuDocument document, string folderPath,
        OfficeImageExportOptions? options = null, DjVuRenderOptions? renderOptions = null, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        OfficeImageExportFormat format = options?.Format ?? OfficeImageExportFormat.Png;
        var operation = CreateOperation(document, options, renderOptions);
        var saved = await operation.SaveFilesAsync(folderPath, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("DjVu", format,
            saved.Report.NativeDiagnostics, saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    private static OfficeImageExportBatchOperation<OfficeImageExportOptions> CreateOperation(DjVuDocument document,
        OfficeImageExportOptions? options, DjVuRenderOptions? renderOptions) {
        ArgumentNullException.ThrowIfNull(document);
        OfficeImageExportOptions snapshot = (options ?? new OfficeImageExportOptions()).CreateEffectiveImageExportOptions<OfficeImageExportOptions>();
        DjVuRenderOptions native = renderOptions?.Clone() ?? new DjVuRenderOptions();
        return OfficeImageExportAdapter.CreateOperation(document.Pages, snapshot, (page, number, effective, token) => {
            return page.ToImage(effective, native, token).WithIdentity("Page " + number, "DjVu");
        });
    }
}
