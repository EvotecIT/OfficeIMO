using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

public static partial class HtmlImageExportExtensions {
    /// <summary>Saves selected images without retaining payloads and returns aggregate conversion evidence.</summary>
    public static OfficeConversionReport SaveAsImages(this HtmlConversionDocument document, string folderPath,
        HtmlRenderOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        HtmlRenderOptions resolved = options?.Clone() ?? new HtmlRenderOptions();
        OfficeImageExportBatchSaveResult saved = CreateHtmlConversionDocumentImageBatchOperation(document, resolved).SaveFiles(folderPath, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("HTML", resolved.Format, saved.Report.Diagnostics,
            saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    /// <summary>Asynchronously saves selected images, retaining paths and evidence rather than encoded payloads.</summary>
    public static async Task<OfficeConversionReport> SaveAsImagesAsync(this HtmlConversionDocument document, string folderPath,
        HtmlRenderOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        HtmlRenderOptions resolved = options?.Clone() ?? new HtmlRenderOptions();
        OfficeImageExportBatchSaveResult saved = await CreateHtmlConversionDocumentImageBatchOperation(document, resolved)
            .SaveFilesAsync(folderPath, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("HTML", resolved.Format, saved.Report.Diagnostics,
            saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    private static OfficeImageExportBatchOperation<HtmlRenderOptions> CreateHtmlConversionDocumentImageBatchOperation(
        HtmlConversionDocument document, HtmlRenderOptions options) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return new OfficeImageExportBatchOperation<HtmlRenderOptions>(options, options.Format, options.FileConflictPolicy,
            (encoding, effective) => document.ExportImages(encoding, effective),
            exportEach: (encoding, effective, consumer, token) => document.ExportImages(encoding, consumer, effective, token),
            exportEachAsync: (encoding, effective, consumer, token) => document.ExportImagesAsync(encoding, consumer, effective, token));
    }
}
