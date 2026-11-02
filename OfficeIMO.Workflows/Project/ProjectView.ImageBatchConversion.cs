using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Project;

namespace OfficeIMO.Workflows;

public static partial class ProjectReportWorkflow {
    /// <summary>Saves selected images without retaining payloads and returns aggregate conversion evidence.</summary>
    public static OfficeConversionReport SaveAsImages(this ProjectView document, string folderPath,
        ProjectImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        ProjectImageExportOptions resolved = options?.Snapshot() ?? new ProjectImageExportOptions();
        OfficeImageExportBatchSaveResult saved = CreateProjectViewImageBatchOperation(document, resolved).SaveFiles(folderPath, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("Project", resolved.Format, saved.Report.Diagnostics,
            saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    /// <summary>Asynchronously saves selected images, retaining paths and evidence rather than encoded payloads.</summary>
    public static async Task<OfficeConversionReport> SaveAsImagesAsync(this ProjectView document, string folderPath,
        ProjectImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        ProjectImageExportOptions resolved = options?.Snapshot() ?? new ProjectImageExportOptions();
        OfficeImageExportBatchSaveResult saved = await CreateProjectViewImageBatchOperation(document, resolved)
            .SaveFilesAsync(folderPath, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("Project", resolved.Format, saved.Report.Diagnostics,
            saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    /// <summary>Asynchronously exports selected images through the shared bounded streaming bridge.</summary>
    public static Task<IReadOnlyList<OfficeImageExportResult>> ExportImagesAsync(this ProjectView document,
        OfficeImageExportFormat format, ProjectImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        ProjectImageExportOptions resolved = options?.Snapshot() ?? new ProjectImageExportOptions();
        resolved.Format = format;
        return CreateProjectViewImageBatchOperation(document, resolved).ExportAsync(cancellationToken);
    }

    /// <summary>Asynchronously delivers selected images with backpressure and cooperative cancellation.</summary>
    public static Task ExportImagesAsync(this ProjectView document, OfficeImageExportFormat format,
        OfficeImageExportAsyncConsumer consumer, ProjectImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        ProjectImageExportOptions resolved = options?.Snapshot() ?? new ProjectImageExportOptions();
        resolved.Format = format;
        return CreateProjectViewImageBatchOperation(document, resolved).ExportEachAsync(consumer, cancellationToken);
    }

    private static OfficeImageExportBatchOperation<ProjectImageExportOptions> CreateProjectViewImageBatchOperation(
        ProjectView document, ProjectImageExportOptions options) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return new OfficeImageExportBatchOperation<ProjectImageExportOptions>(options, options.Format, options.FileConflictPolicy,
            (encoding, effective) => document.ExportImages(encoding, effective),
            exportEach: (encoding, effective, consumer, token) => document.ExportImages(encoding, consumer, effective, token));
    }
}
