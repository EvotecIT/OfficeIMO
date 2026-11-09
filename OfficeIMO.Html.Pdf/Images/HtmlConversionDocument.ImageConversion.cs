using System.Diagnostics;
using OfficeIMO.Drawing;

using OfficeIMO.Html;

namespace OfficeIMO.Html.Pdf;

public static partial class HtmlImageExportExtensions {
    /// <summary>Renders the selected page using the encoding in <see cref="OfficeImageExportOptions.Format"/>.</summary>
    public static OfficeImageExportResult ToImage(this HtmlConversionDocument document, HtmlRenderOptions? options = null,
        CancellationToken cancellationToken = default) => document.ToImage(options, out _, cancellationToken);

    /// <summary>Renders the selected page and returns its immutable fidelity report.</summary>
    public static OfficeImageExportResult ToImage(this HtmlConversionDocument document, HtmlRenderOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        HtmlRenderOptions resolved = SnapshotHtmlConversionDocumentImageOptions(options);
        OfficeImageExportResult result = CreateImageExportOperation(document, resolved).Export(cancellationToken);
        report = OfficeImageExportEvidence.CreateReport("HTML", result, timer.Elapsed);
        return result;
    }

    /// <summary>Renders the selected page with cancellation and returns the encoded image and its diagnostics.</summary>
    public static async Task<OfficeImageExportResult> ToImageAsync(this HtmlConversionDocument document,
        HtmlRenderOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        HtmlRenderOptions resolved = SnapshotHtmlConversionDocumentImageOptions(options);
        OfficeImageExportResult result = await CreateImageExportOperation(document, resolved)
            .ExportAsync(cancellationToken).ConfigureAwait(false);
        return result;
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
        ArgumentNullException.ThrowIfNull(document);
        return new OfficeImageExportOperation<HtmlRenderOptions>(options, options.Format, options.FileConflictPolicy,
            (encoding, effective) => document.ToImage(encoding, effective, effective.PageIndex),
            exportWithCancellation: (encoding, effective, token) => document.ToImage(encoding, effective, effective.PageIndex, token),
            exportAsync: (encoding, effective, token) => document.ToImageAsync(encoding, effective, effective.PageIndex, token));
    }
}
