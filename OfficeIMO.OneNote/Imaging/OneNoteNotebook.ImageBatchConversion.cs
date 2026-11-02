using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.OneNote;

public static partial class OneNoteImageExportExtensions {
    /// <summary>Saves selected images without retaining payloads and returns aggregate conversion evidence.</summary>
    public static OfficeConversionReport SaveAsImages(this OneNoteNotebook document, string folderPath,
        OneNotePageBatchRenderingOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        OneNotePageBatchRenderingOptions resolved = options?.CloneBatch() ?? new OneNotePageBatchRenderingOptions();
        OfficeImageExportBatchSaveResult saved = CreateOneNoteNotebookImageBatchOperation(document, resolved).SaveFiles(folderPath, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("OneNote", resolved.Format, saved.Report.Diagnostics,
            saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    /// <summary>Asynchronously saves selected images, retaining paths and evidence rather than encoded payloads.</summary>
    public static async Task<OfficeConversionReport> SaveAsImagesAsync(this OneNoteNotebook document, string folderPath,
        OneNotePageBatchRenderingOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        OneNotePageBatchRenderingOptions resolved = options?.CloneBatch() ?? new OneNotePageBatchRenderingOptions();
        OfficeImageExportBatchSaveResult saved = await CreateOneNoteNotebookImageBatchOperation(document, resolved)
            .SaveFilesAsync(folderPath, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("OneNote", resolved.Format, saved.Report.Diagnostics,
            saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    /// <summary>Asynchronously exports selected images through the shared bounded streaming bridge.</summary>
    public static Task<IReadOnlyList<OfficeImageExportResult>> ExportImagesAsync(this OneNoteNotebook document,
        OfficeImageExportFormat format, OneNotePageBatchRenderingOptions? options = null, CancellationToken cancellationToken = default) {
        OneNotePageBatchRenderingOptions resolved = options?.CloneBatch() ?? new OneNotePageBatchRenderingOptions();
        resolved.Format = format;
        return CreateOneNoteNotebookImageBatchOperation(document, resolved).ExportAsync(cancellationToken);
    }

    /// <summary>Asynchronously delivers selected images with backpressure and cooperative cancellation.</summary>
    public static Task ExportImagesAsync(this OneNoteNotebook document, OfficeImageExportFormat format,
        OfficeImageExportAsyncConsumer consumer, OneNotePageBatchRenderingOptions? options = null, CancellationToken cancellationToken = default) {
        OneNotePageBatchRenderingOptions resolved = options?.CloneBatch() ?? new OneNotePageBatchRenderingOptions();
        resolved.Format = format;
        return CreateOneNoteNotebookImageBatchOperation(document, resolved).ExportEachAsync(consumer, cancellationToken);
    }

    /// <summary>Streams selected images under the shared deadline, admission and aggregate resource limits.</summary>
    public static void ExportImages(this OneNoteNotebook document, OfficeImageExportFormat format,
        OfficeImageExportConsumer consumer, OneNotePageBatchRenderingOptions? options = null, CancellationToken cancellationToken = default) {
        OneNotePageBatchRenderingOptions resolved = options?.CloneBatch() ?? new OneNotePageBatchRenderingOptions();
        resolved.Format = format;
        CreateOneNoteNotebookImageBatchOperation(document, resolved).ExportEach(consumer, cancellationToken);
    }

    private static OfficeImageExportBatchOperation<OneNotePageBatchRenderingOptions> CreateOneNoteNotebookImageBatchOperation(
        OneNoteNotebook document, OneNotePageBatchRenderingOptions options) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return new OfficeImageExportBatchOperation<OneNotePageBatchRenderingOptions>(options, options.Format, options.FileConflictPolicy,
            (encoding, effective) => OneNoteImageExportEngine.ExportNotebook(document, encoding, effective),
            exportEach: (encoding, effective, consumer, token) => OneNoteImageExportEngine.ExportNotebook(document, encoding, effective, consumer, token));
    }
}
