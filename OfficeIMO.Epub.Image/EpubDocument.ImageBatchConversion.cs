using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.Epub.Image;

public static partial class EpubImageExportExtensions {
    /// <summary>Saves selected images without retaining payloads and returns aggregate conversion evidence.</summary>
    public static OfficeConversionReport SaveAsImages(this EpubDocument document, string folderPath,
        EpubImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        EpubImageExportOptions resolved = options?.CloneEpub() ?? new EpubImageExportOptions();
        OfficeImageExportBatchSaveResult saved = CreateEpubDocumentImageBatchOperation(document, resolved).SaveFiles(folderPath, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("EPUB", resolved.Format, saved.Report.Diagnostics,
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
        return OfficeImageExportEvidence.CreateReport("EPUB", resolved.Format, saved.Report.Diagnostics,
            saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    private static OfficeImageExportBatchOperation<EpubImageExportOptions> CreateEpubDocumentImageBatchOperation(
        EpubDocument document, EpubImageExportOptions options) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return new OfficeImageExportBatchOperation<EpubImageExportOptions>(options, options.Format, options.FileConflictPolicy,
            (encoding, effective) => document.ExportImages(encoding, effective),
            exportEach: (encoding, effective, consumer, token) => document.ExportImages(encoding, consumer, effective, token),
            exportEachAsync: (encoding, effective, consumer, token) => document.ExportImagesAsync(encoding, consumer, effective, token));
    }
}
