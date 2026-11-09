using System.Threading.Tasks;
using System.Diagnostics;
using OfficeIMO.Drawing;

namespace OfficeIMO.Xps;

/// <summary>Renders XPS surfaces through the existing native projection and shared image policies.</summary>
public static class XpsDocumentImageConversionExtensions {
    /// <summary>Renders selected source pages in caller order, retaining source and encoding diagnostics.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ToImages(this XpsDocument document,
        OfficeImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
        CreateOperation(document, options).Export(cancellationToken);

    /// <summary>Streams selected images without retaining earlier encoded payloads.</summary>
    public static void ToImages(this XpsDocument document, OfficeImageExportConsumer consumer,
        OfficeImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
        CreateOperation(document, options).ExportEach(consumer, cancellationToken);

    /// <summary>Renders selected source pages asynchronously with shared cancellation and deadlines.</summary>
    public static Task<IReadOnlyList<OfficeImageExportResult>> ToImagesAsync(this XpsDocument document,
        OfficeImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
        CreateOperation(document, options).ExportAsync(cancellationToken);

    /// <summary>Saves selected pages with atomic files, bounded output and fidelity acceptance.</summary>
    public static OfficeConversionReport SaveAsImages(this XpsDocument document, string folderPath,
        OfficeImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        OfficeImageExportFormat format = options?.Format ?? OfficeImageExportFormat.Png;
        var operation = CreateOperation(document, options);
        var saved = operation.SaveFiles(folderPath, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("XPS", format,
            saved.Report.NativeDiagnostics, saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    /// <summary>Saves selected pages asynchronously without retaining completed image payloads.</summary>
    public static async Task<OfficeConversionReport> SaveAsImagesAsync(this XpsDocument document, string folderPath,
        OfficeImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        OfficeImageExportFormat format = options?.Format ?? OfficeImageExportFormat.Png;
        var operation = CreateOperation(document, options);
        var saved = await operation.SaveFilesAsync(folderPath, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("XPS", format,
            saved.Report.NativeDiagnostics, saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    private static OfficeImageExportBatchOperation<OfficeImageExportOptions> CreateOperation(XpsDocument document,
        OfficeImageExportOptions? options) {
        ArgumentNullException.ThrowIfNull(document);
        OfficeImageExportOptions snapshot = (options ?? new OfficeImageExportOptions()).CreateEffectiveImageExportOptions<OfficeImageExportOptions>();
        return OfficeImageExportAdapter.CreateOperation(document.Pages, snapshot, (page, number, effective, token) => {
            return page.ToImage(effective.Format, effective, token).WithIdentity("Page " + number, "XPS");
        });
    }
}
