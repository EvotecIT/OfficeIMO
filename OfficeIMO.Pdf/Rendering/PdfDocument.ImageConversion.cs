using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public static partial class PdfImageExportExtensions {
    /// <summary>Renders the selected PDF pages using the encoding and resource limits in the options.</summary>
    public static IReadOnlyList<byte[]> ExportImages(this PdfDocument document,
        PdfImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
        document.ExportImages(options, out _, cancellationToken);

    /// <summary>Renders the selected pages and returns an immutable aggregate fidelity report.</summary>
    public static IReadOnlyList<byte[]> ExportImages(this PdfDocument document, PdfImageExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        PdfImageExportOptions snapshot = options?.Clone() ?? new PdfImageExportOptions();
        var images = new List<OfficeImageExportResult>();
        document.CreateImageExportOperation(snapshot).ExportEach(images.Add, cancellationToken);
        report = OfficeImageExportEvidence.CreateReport("Pdf", snapshot.Format,
            images.SelectMany(image => image.Diagnostics), images.Count, timer.Elapsed);
        return images.Select(image => image.Bytes).ToList().AsReadOnly();
    }

    /// <summary>Asynchronously renders selected pages with a cloned options snapshot and fidelity report.</summary>
    public static async Task<OfficeConversionResult<IReadOnlyList<byte[]>>> ExportImagesAsync(this PdfDocument document,
        PdfImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        PdfImageExportOptions snapshot = options?.Clone() ?? new PdfImageExportOptions();
        IReadOnlyList<OfficeImageExportResult> images = await document.CreateImageExportOperation(snapshot)
            .ExportAsync(cancellationToken).ConfigureAwait(false);
        return new OfficeConversionResult<IReadOnlyList<byte[]>>(
            images.Select(image => image.Bytes).ToList().AsReadOnly(),
            OfficeImageExportEvidence.CreateReport("Pdf", snapshot.Format,
                images.SelectMany(image => image.Diagnostics), images.Count, timer.Elapsed));
    }

    /// <summary>Saves selected pages without retaining their encoded payloads and returns the fidelity report.</summary>
    public static OfficeConversionReport SaveAsImages(this PdfDocument document, string folderPath,
        PdfImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        PdfImageExportOptions snapshot = options?.Clone() ?? new PdfImageExportOptions();
        OfficeImageExportBatchSaveResult saved = document.CreateImageExportOperation(snapshot).SaveFiles(folderPath, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("Pdf", snapshot.Format, saved.Report.Diagnostics,
            saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    /// <summary>Asynchronously saves selected pages without retaining their encoded payloads.</summary>
    public static async Task<OfficeConversionReport> SaveAsImagesAsync(this PdfDocument document, string folderPath,
        PdfImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        PdfImageExportOptions snapshot = options?.Clone() ?? new PdfImageExportOptions();
        OfficeImageExportBatchSaveResult saved = await document.CreateImageExportOperation(snapshot)
            .SaveFilesAsync(folderPath, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("Pdf", snapshot.Format, saved.Report.Diagnostics,
            saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }
}
