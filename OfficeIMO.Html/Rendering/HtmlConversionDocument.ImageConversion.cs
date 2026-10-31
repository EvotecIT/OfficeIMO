using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

public static partial class HtmlImageExportExtensions {
    /// <summary>Renders the selected page using the encoding in <see cref="OfficeImageExportOptions.Format"/>.</summary>
    public static byte[] ToImage(this HtmlConversionDocument document, HtmlRenderOptions? options = null,
        CancellationToken cancellationToken = default) => document.ToImage(options, out _, cancellationToken);

    /// <summary>Renders the selected page and returns its immutable fidelity report.</summary>
    public static byte[] ToImage(this HtmlConversionDocument document, HtmlRenderOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        HtmlRenderOptions resolved = SnapshotHtmlConversionDocumentImageOptions(options);
        OfficeImageExportResult result = CreateImageExportOperation(document, resolved).Export(cancellationToken);
        report = OfficeImageExportEvidence.CreateReport("HTML", result, timer.Elapsed);
        return result.Bytes;
    }

    /// <summary>Renders the selected page with cancellation and returns the bytes and fidelity report.</summary>
    public static async Task<OfficeConversionResult<byte[]>> ToImageAsync(this HtmlConversionDocument document,
        HtmlRenderOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        HtmlRenderOptions resolved = SnapshotHtmlConversionDocumentImageOptions(options);
        OfficeImageExportResult result = await CreateImageExportOperation(document, resolved)
            .ExportAsync(cancellationToken).ConfigureAwait(false);
        return new OfficeConversionResult<byte[]>(result.Bytes, OfficeImageExportEvidence.CreateReport("HTML", result, timer.Elapsed));
    }

    /// <summary>Writes the selected page to a file and returns its fidelity report.</summary>
    public static OfficeConversionReport SaveAsImage(this HtmlConversionDocument document, string path, HtmlRenderOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        HtmlRenderOptions resolved = SnapshotHtmlConversionDocumentImageOptions(options);
        OfficeImageExportResult result = CreateImageExportOperation(document, resolved).Save(path, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("HTML", result, timer.Elapsed);
    }

    /// <summary>Writes the selected page to a caller-owned stream and returns its fidelity report.</summary>
    public static OfficeConversionReport SaveAsImage(this HtmlConversionDocument document, Stream stream, HtmlRenderOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        HtmlRenderOptions resolved = SnapshotHtmlConversionDocumentImageOptions(options);
        OfficeImageExportResult result = CreateImageExportOperation(document, resolved).Save(stream, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("HTML", result, timer.Elapsed);
    }

    /// <summary>Asynchronously writes the selected page to a file and returns its fidelity report.</summary>
    public static async Task<OfficeConversionReport> SaveAsImageAsync(this HtmlConversionDocument document, string path,
        HtmlRenderOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        HtmlRenderOptions resolved = SnapshotHtmlConversionDocumentImageOptions(options);
        OfficeImageExportResult result = await CreateImageExportOperation(document, resolved)
            .SaveAsync(path, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("HTML", result, timer.Elapsed);
    }

    /// <summary>Asynchronously writes the selected page to a caller-owned stream and returns its fidelity report.</summary>
    public static async Task<OfficeConversionReport> SaveAsImageAsync(this HtmlConversionDocument document, Stream stream,
        HtmlRenderOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        HtmlRenderOptions resolved = SnapshotHtmlConversionDocumentImageOptions(options);
        OfficeImageExportResult result = await CreateImageExportOperation(document, resolved)
            .SaveAsync(stream, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("HTML", result, timer.Elapsed);
    }

    private static HtmlRenderOptions SnapshotHtmlConversionDocumentImageOptions(HtmlRenderOptions? options) =>
        options?.Clone() ?? new HtmlRenderOptions();

    private static OfficeImageExportOperation<HtmlRenderOptions> CreateImageExportOperation(
        HtmlConversionDocument document, HtmlRenderOptions options) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return new OfficeImageExportOperation<HtmlRenderOptions>(options, options.Format, options.FileConflictPolicy,
            (encoding, effective) => document.ExportImage(encoding, effective, effective.PageIndex),
            exportWithCancellation: (encoding, effective, token) => document.ExportImage(encoding, effective, effective.PageIndex, token),
            exportAsync: (encoding, effective, token) => document.ExportImageAsync(encoding, effective, effective.PageIndex, token));
    }
    private static OfficeImageExportOperation<HtmlRenderOptions> CreateImageExportOperation(
        HtmlConversionDocument document, HtmlRenderOptions? options, OfficeImageExportFormat format, int pageIndex) {
        if (pageIndex < 0) throw new ArgumentOutOfRangeException(nameof(pageIndex));
        HtmlRenderOptions resolved = SnapshotHtmlConversionDocumentImageOptions(options);
        resolved.Format = format;
        resolved.PageIndex = pageIndex;
        return CreateImageExportOperation(document, resolved);
    }

}
