using System.Diagnostics;
using OfficeIMO.Drawing;

namespace OfficeIMO.Epub.Image;

public static partial class EpubImageExportExtensions {
    /// <summary>Renders a batch using the encoding and selection in the options.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ToImages(this EpubDocument document, EpubImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        EpubImageExportOptions resolved = options?.CloneEpub() ?? new EpubImageExportOptions();
        return CreateEpubDocumentImageBatchOperation(document, resolved).Export(cancellationToken);
    }

    /// <summary>Renders a batch asynchronously using the same options and evidence contract.</summary>
    public static Task<IReadOnlyList<OfficeImageExportResult>> ToImagesAsync(this EpubDocument document, EpubImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        EpubImageExportOptions resolved = options?.CloneEpub() ?? new EpubImageExportOptions();
        return CreateEpubDocumentImageBatchOperation(document, resolved).ExportAsync(cancellationToken);
    }

    /// <summary>Saves selected images without retaining payloads and returns aggregate conversion evidence.</summary>
    public static OfficeConversionReport SaveAsImages(this EpubDocument document, string folderPath,
        EpubImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        EpubImageExportOptions resolved = options?.CloneEpub() ?? new EpubImageExportOptions();
        OfficeImageExportBatchSaveResult saved = CreateEpubDocumentImageBatchOperation(document, resolved).SaveFiles(folderPath, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("EPUB", resolved.Format, saved.Report.NativeDiagnostics,
            saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    /// <summary>Asynchronously saves selected images, retaining paths and evidence rather than encoded payloads.</summary>
    public static async Task<OfficeConversionReport> SaveAsImagesAsync(this EpubDocument document, string folderPath,
        EpubImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        EpubImageExportOptions resolved = options?.CloneEpub() ?? new EpubImageExportOptions();
        OfficeImageExportBatchSaveResult saved = await CreateEpubDocumentImageBatchOperation(document, resolved)
            .SaveFilesAsync(folderPath, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("EPUB", resolved.Format, saved.Report.NativeDiagnostics,
            saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    private static OfficeImageExportBatchOperation<EpubImageExportOptions> CreateEpubDocumentImageBatchOperation(
        EpubDocument document, EpubImageExportOptions options) {
        ArgumentNullException.ThrowIfNull(document);
        return new OfficeImageExportBatchOperation<EpubImageExportOptions>(options, options.Format, options.FileConflictPolicy,
            (encoding, effective) => document.ToImages(encoding, effective),
            exportEach: (encoding, effective, consumer, token) => document.ToImages(encoding, consumer, effective, token),
            exportEachAsync: (encoding, effective, consumer, token) => document.ToImagesAsync(encoding, consumer, effective, token));
    }
}
