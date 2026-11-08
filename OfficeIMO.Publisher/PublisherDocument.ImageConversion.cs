using System.Threading.Tasks;
using System.Diagnostics;
using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher;

/// <summary>Renders Publisher surfaces through the existing native projection and shared image policies.</summary>
public static class PublisherDocumentImageConversionExtensions {
    /// <summary>Renders selected source pages in caller order, retaining source and encoding diagnostics.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ToImages(this PublisherDocument document,
        OfficeImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
        CreateOperation(document, options).Export(cancellationToken);

    /// <summary>Streams selected images without retaining earlier encoded payloads.</summary>
    public static void ToImages(this PublisherDocument document, OfficeImageExportConsumer consumer,
        OfficeImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
        CreateOperation(document, options).ExportEach(consumer, cancellationToken);

    /// <summary>Renders selected source pages asynchronously with shared cancellation and deadlines.</summary>
    public static Task<IReadOnlyList<OfficeImageExportResult>> ToImagesAsync(this PublisherDocument document,
        OfficeImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
        CreateOperation(document, options).ExportAsync(cancellationToken);

    /// <summary>Saves selected pages with atomic files, bounded output and fidelity acceptance.</summary>
    public static OfficeConversionReport SaveAsImages(this PublisherDocument document, string folderPath,
        OfficeImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        OfficeImageExportFormat format = options?.Format ?? OfficeImageExportFormat.Png;
        var operation = CreateOperation(document, options);
        var saved = operation.SaveFiles(folderPath, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("Publisher", format,
            saved.Report.NativeDiagnostics, saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    /// <summary>Saves selected pages asynchronously without retaining completed image payloads.</summary>
    public static async Task<OfficeConversionReport> SaveAsImagesAsync(this PublisherDocument document, string folderPath,
        OfficeImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        OfficeImageExportFormat format = options?.Format ?? OfficeImageExportFormat.Png;
        var operation = CreateOperation(document, options);
        var saved = await operation.SaveFilesAsync(folderPath, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("Publisher", format,
            saved.Report.NativeDiagnostics, saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    private static OfficeImageExportBatchOperation<OfficeImageExportOptions> CreateOperation(PublisherDocument document,
        OfficeImageExportOptions? options) {
        ArgumentNullException.ThrowIfNull(document);
        OfficeImageExportOptions snapshot = (options ?? new OfficeImageExportOptions()).SnapshotForUnits(72D);
        return OfficeImageExportAdapter.CreateOperation(document.Pages, snapshot, (page, number, effective, token) => {
            return effective.EnsureAccepted(OfficeImageExportEvidence.AttachSourceReport(page.CreateDrawing().ToImage(effective.Format, effective, token).WithIdentity(page.Name, "Publisher"), document.ReadReport));
        });
    }
}
