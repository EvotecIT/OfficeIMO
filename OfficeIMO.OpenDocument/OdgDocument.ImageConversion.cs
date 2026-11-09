using System.Threading.Tasks;
using System.Diagnostics;
using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

/// <summary>Renders ODG surfaces through the existing native projection and shared image policies.</summary>
public static class OdgDocumentImageConversionExtensions {
    /// <summary>Renders selected source pages in caller order, retaining source and encoding diagnostics.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ToImages(this OdgDocument document,
        OfficeImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
        CreateOperation(document, options).Export(cancellationToken);

    /// <summary>Streams selected images without retaining earlier encoded payloads.</summary>
    public static void ToImages(this OdgDocument document, OfficeImageExportConsumer consumer,
        OfficeImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
        CreateOperation(document, options).ExportEach(consumer, cancellationToken);

    /// <summary>Renders selected source pages asynchronously with shared cancellation and deadlines.</summary>
    public static Task<IReadOnlyList<OfficeImageExportResult>> ToImagesAsync(this OdgDocument document,
        OfficeImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
        CreateOperation(document, options).ExportAsync(cancellationToken);

    /// <summary>Saves selected pages with atomic files, bounded output and fidelity acceptance.</summary>
    public static OfficeConversionReport SaveAsImages(this OdgDocument document, string folderPath,
        OfficeImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        OfficeImageExportFormat format = options?.Format ?? OfficeImageExportFormat.Png;
        var operation = CreateOperation(document, options);
        var saved = operation.SaveFiles(folderPath, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("ODG", format,
            saved.Report.NativeDiagnostics, saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    /// <summary>Saves selected pages asynchronously without retaining completed image payloads.</summary>
    public static async Task<OfficeConversionReport> SaveAsImagesAsync(this OdgDocument document, string folderPath,
        OfficeImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        OfficeImageExportFormat format = options?.Format ?? OfficeImageExportFormat.Png;
        var operation = CreateOperation(document, options);
        var saved = await operation.SaveFilesAsync(folderPath, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("ODG", format,
            saved.Report.NativeDiagnostics, saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    private static OfficeImageExportBatchOperation<OfficeImageExportOptions> CreateOperation(OdgDocument document,
        OfficeImageExportOptions? options) {
        ArgumentNullException.ThrowIfNull(document);
        OfficeImageExportOptions snapshot = (options ?? new OfficeImageExportOptions()).SnapshotForUnits(72D);
        return OfficeImageExportAdapter.CreateOperation(document.Pages, snapshot, (page, number, effective, token) => {
            var profile = new OfficeRenderingProfile("ODG image", effective.Fonts, effective.TextShapingProvider, effective.TextShapingLanguage, effective.ImageCodec, effective.Policy);
            var drawing = page.ToDrawing(profile, cancellationToken: token);
            return effective.EnsureAccepted(OfficeImageExportEvidence.AttachSourceReport(drawing.Value.ToImage(effective.Format, effective, token).WithIdentity("Page " + number, "ODG"), drawing.Report));
        });
    }
}
