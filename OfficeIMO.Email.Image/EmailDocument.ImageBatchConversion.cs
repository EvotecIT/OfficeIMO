using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.Email;

public static partial class EmailImageExportExtensions {
    /// <summary>Saves selected images without retaining payloads and returns aggregate conversion evidence.</summary>
    public static OfficeConversionReport SaveAsImages(this EmailDocument document, string folderPath,
        EmailImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        EmailImageExportOptions resolved = options?.CloneEmail() ?? new EmailImageExportOptions();
        OfficeImageExportBatchSaveResult saved = CreateEmailDocumentImageBatchOperation(document, resolved).SaveFiles(folderPath, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("Email", resolved.Format, saved.Report.Diagnostics,
            saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    /// <summary>Asynchronously saves selected images, retaining paths and evidence rather than encoded payloads.</summary>
    public static async Task<OfficeConversionReport> SaveAsImagesAsync(this EmailDocument document, string folderPath,
        EmailImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        EmailImageExportOptions resolved = options?.CloneEmail() ?? new EmailImageExportOptions();
        OfficeImageExportBatchSaveResult saved = await CreateEmailDocumentImageBatchOperation(document, resolved)
            .SaveFilesAsync(folderPath, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("Email", resolved.Format, saved.Report.Diagnostics,
            saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    private static OfficeImageExportBatchOperation<EmailImageExportOptions> CreateEmailDocumentImageBatchOperation(
        EmailDocument document, EmailImageExportOptions options) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return new OfficeImageExportBatchOperation<EmailImageExportOptions>(options, options.Format, options.FileConflictPolicy,
            (encoding, effective) => document.ExportImages(encoding, effective),
            exportEach: (encoding, effective, consumer, token) => document.ExportImages(encoding, consumer, effective, token),
            exportEachAsync: (encoding, effective, consumer, token) => document.ExportImagesAsync(encoding, consumer, effective, token));
    }
}
