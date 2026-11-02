using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

public static partial class VisioImageExportExtensions {
    /// <summary>Saves selected images without retaining payloads and returns aggregate conversion evidence.</summary>
    public static OfficeConversionReport SaveAsImages(this VisioDocument document, string folderPath,
        VisioImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        VisioImageExportOptions resolved = options?.Clone() ?? new VisioImageExportOptions();
        OfficeImageExportBatchSaveResult saved = CreateVisioDocumentImageBatchOperation(document, resolved).SaveFiles(folderPath, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("Visio", resolved.Format, saved.Report.Diagnostics,
            saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    /// <summary>Asynchronously saves selected images, retaining paths and evidence rather than encoded payloads.</summary>
    public static async Task<OfficeConversionReport> SaveAsImagesAsync(this VisioDocument document, string folderPath,
        VisioImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        VisioImageExportOptions resolved = options?.Clone() ?? new VisioImageExportOptions();
        OfficeImageExportBatchSaveResult saved = await CreateVisioDocumentImageBatchOperation(document, resolved)
            .SaveFilesAsync(folderPath, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("Visio", resolved.Format, saved.Report.Diagnostics,
            saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    /// <summary>Asynchronously exports selected images through the shared bounded streaming bridge.</summary>
    public static Task<IReadOnlyList<OfficeImageExportResult>> ExportImagesAsync(this VisioDocument document,
        OfficeImageExportFormat format, VisioImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        VisioImageExportOptions resolved = options?.Clone() ?? new VisioImageExportOptions();
        resolved.Format = format;
        return CreateVisioDocumentImageBatchOperation(document, resolved).ExportAsync(cancellationToken);
    }

    /// <summary>Asynchronously delivers selected images with backpressure and cooperative cancellation.</summary>
    public static Task ExportImagesAsync(this VisioDocument document, OfficeImageExportFormat format,
        OfficeImageExportAsyncConsumer consumer, VisioImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        VisioImageExportOptions resolved = options?.Clone() ?? new VisioImageExportOptions();
        resolved.Format = format;
        return CreateVisioDocumentImageBatchOperation(document, resolved).ExportEachAsync(consumer, cancellationToken);
    }

    /// <summary>Saves images with per-file metadata for the existing in-process host contract.</summary>
    internal static IReadOnlyList<OfficeImageExportResult> SaveImageResults(this VisioDocument document,
        string folderPath, VisioImageExportOptions options) =>
        CreateVisioDocumentImageBatchOperation(document, options).Save(folderPath);

    private static OfficeImageExportBatchOperation<VisioImageExportOptions> CreateVisioDocumentImageBatchOperation(
        VisioDocument document, VisioImageExportOptions options) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return new OfficeImageExportBatchOperation<VisioImageExportOptions>(options, options.Format, options.FileConflictPolicy,
            (encoding, effective) => document.ExportImages(encoding, effective),
            exportEach: (encoding, effective, consumer, token) => document.ExportImages(encoding, consumer, effective, token));
    }
}
