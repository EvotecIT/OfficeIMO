using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

using OfficeIMO.Html;

namespace OfficeIMO.Html.Pdf;

public static partial class HtmlImageExportExtensions {
    /// <summary>Renders a batch using the encoding and selection in the options.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ToImages(this HtmlConversionDocument document, HtmlRenderOptions? options = null,
        CancellationToken cancellationToken = default) {
        HtmlRenderOptions resolved = options?.Clone() ?? new HtmlRenderOptions();
        return CreateHtmlConversionDocumentImageBatchOperation(document, resolved).Export(cancellationToken);
    }

    /// <summary>Renders a batch asynchronously using the same options and evidence contract.</summary>
    public static Task<IReadOnlyList<OfficeImageExportResult>> ToImagesAsync(this HtmlConversionDocument document, HtmlRenderOptions? options = null,
        CancellationToken cancellationToken = default) {
        HtmlRenderOptions resolved = options?.Clone() ?? new HtmlRenderOptions();
        return CreateHtmlConversionDocumentImageBatchOperation(document, resolved).ExportAsync(cancellationToken);
    }

    /// <summary>Saves selected images without retaining payloads and returns aggregate conversion evidence.</summary>
    public static OfficeConversionReport SaveAsImages(this HtmlConversionDocument document, string folderPath,
        HtmlRenderOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        HtmlRenderOptions resolved = options?.Clone() ?? new HtmlRenderOptions();
        OfficeImageExportBatchSaveResult saved = CreateHtmlConversionDocumentImageBatchOperation(document, resolved).SaveFiles(folderPath, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("HTML", resolved.Format, saved.Report.NativeDiagnostics,
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
        return OfficeImageExportEvidence.CreateReport("HTML", resolved.Format, saved.Report.NativeDiagnostics,
            saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    private static OfficeImageExportBatchOperation<HtmlRenderOptions> CreateHtmlConversionDocumentImageBatchOperation(
        HtmlConversionDocument document, HtmlRenderOptions options) {
        ArgumentNullException.ThrowIfNull(document);
        return new OfficeImageExportBatchOperation<HtmlRenderOptions>(options, options.Format, options.FileConflictPolicy,
            (encoding, effective) => document.ToImages(encoding, effective),
            exportEach: (encoding, effective, consumer, token) => document.ToImages(encoding, consumer, effective, token),
            exportEachAsync: (encoding, effective, consumer, token) => document.ToImagesAsync(encoding, consumer, effective, token));
    }
}
