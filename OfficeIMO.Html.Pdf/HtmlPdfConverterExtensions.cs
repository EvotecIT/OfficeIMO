using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

/// <summary>Direct HTML-to-PDF conversion helpers backed by the shared OfficeIMO HTML render scene.</summary>
public static partial class HtmlPdfConverterExtensions {
    /// <summary>Converts HTML to PDF bytes using the same layout engine as HTML PNG and SVG export.</summary>
    /// <example><code>byte[] pdf = html.ToPdf();</code></example>
    public static byte[] ToPdf(this string html, HtmlPdfSaveOptions? options = null) =>
        html.ToPdfDocumentResult(options).ToBytes();

    /// <summary>Converts a shared HTML conversion document to PDF bytes.</summary>
    public static byte[] ToPdf(this HtmlConversionDocument document, HtmlPdfSaveOptions? options = null) =>
        document.ToPdfDocumentResult(options).ToBytes();

    /// <summary>Reads UTF-8 HTML from a stream and converts it to PDF bytes.</summary>
    public static byte[] ToPdf(this Stream htmlStream, HtmlPdfSaveOptions? options = null) =>
        htmlStream.ToPdfDocumentResult(options).ToBytes();

    /// <summary>Asynchronously resolves HTML resources and converts HTML to PDF bytes.</summary>
    /// <example><code>byte[] pdf = await html.ToPdfAsync(cancellationToken: token);</code></example>
    public static async Task<byte[]> ToPdfAsync(this string html, HtmlPdfSaveOptions? options = null, CancellationToken cancellationToken = default) =>
        SerializeToBytes(await html.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false), cancellationToken);

    /// <summary>Asynchronously converts a shared HTML conversion document to PDF bytes.</summary>
    public static async Task<byte[]> ToPdfAsync(this HtmlConversionDocument document, HtmlPdfSaveOptions? options = null, CancellationToken cancellationToken = default) =>
        SerializeToBytes(await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false), cancellationToken);

    /// <summary>Asynchronously reads UTF-8 HTML from a stream and converts it to PDF bytes.</summary>
    public static async Task<byte[]> ToPdfAsync(this Stream htmlStream, HtmlPdfSaveOptions? options = null, CancellationToken cancellationToken = default) =>
        SerializeToBytes(await htmlStream.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false), cancellationToken);

    /// <summary>Converts HTML to the first-party PDF document model.</summary>
    public static PdfCore.PdfDocument ToPdfDocument(this string html, HtmlPdfSaveOptions? options = null) =>
        html.ToPdfDocumentResult(options).Value;

    /// <summary>Converts a shared HTML conversion document to the first-party PDF document model.</summary>
    public static PdfCore.PdfDocument ToPdfDocument(this HtmlConversionDocument document, HtmlPdfSaveOptions? options = null) =>
        document.ToPdfDocumentResult(options).Value;

    /// <summary>Reads UTF-8 HTML from a stream and converts it to the first-party PDF document model.</summary>
    public static PdfCore.PdfDocument ToPdfDocument(this Stream htmlStream, HtmlPdfSaveOptions? options = null) =>
        htmlStream.ToPdfDocumentResult(options).Value;

    /// <summary>Asynchronously converts HTML to the first-party PDF document model.</summary>
    public static async Task<PdfCore.PdfDocument> ToPdfDocumentAsync(this string html, HtmlPdfSaveOptions? options = null, CancellationToken cancellationToken = default) =>
        (await html.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false)).Value;

    /// <summary>Asynchronously converts a shared HTML conversion document to the first-party PDF document model.</summary>
    public static async Task<PdfCore.PdfDocument> ToPdfDocumentAsync(this HtmlConversionDocument document, HtmlPdfSaveOptions? options = null, CancellationToken cancellationToken = default) =>
        (await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false)).Value;

    /// <summary>Asynchronously reads UTF-8 HTML from a stream and converts it to the first-party PDF document model.</summary>
    public static async Task<PdfCore.PdfDocument> ToPdfDocumentAsync(this Stream htmlStream, HtmlPdfSaveOptions? options = null, CancellationToken cancellationToken = default) =>
        (await htmlStream.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false)).Value;

    /// <summary>Converts HTML to a PDF document plus an immutable diagnostics snapshot.</summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this string html, HtmlPdfSaveOptions? options = null) {
        if (html == null) throw new ArgumentNullException(nameof(html));
        HtmlPdfRenderResult rendered = HtmlPdfRenderedConverter.Convert(html, Normalize(options));
        return CreateResult(rendered);
    }

    /// <summary>Converts a shared HTML conversion document to a PDF document plus diagnostics.</summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this HtmlConversionDocument document, HtmlPdfSaveOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        HtmlPdfRenderResult rendered = HtmlPdfRenderedConverter.Convert(
            document.CreateDocumentForConversion(HtmlCssMediaContext.Print),
            Normalize(options));
        return CreateResult(rendered);
    }

    /// <summary>Reads UTF-8 HTML from a stream and converts it to a PDF document plus diagnostics.</summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this Stream htmlStream, HtmlPdfSaveOptions? options = null) {
        if (htmlStream == null) throw new ArgumentNullException(nameof(htmlStream));
        return HtmlTextIO.Read(htmlStream).ToPdfDocumentResult(options);
    }

    /// <summary>Asynchronously converts HTML to a PDF document plus an immutable diagnostics snapshot.</summary>
    public static async Task<PdfCore.PdfDocumentConversionResult> ToPdfDocumentResultAsync(this string html, HtmlPdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        if (html == null) throw new ArgumentNullException(nameof(html));
        cancellationToken.ThrowIfCancellationRequested();
        HtmlPdfRenderResult rendered = await HtmlPdfRenderedConverter.ConvertAsync(html, Normalize(options), cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return CreateResult(rendered);
    }

    /// <summary>Asynchronously converts a shared HTML conversion document to a PDF document plus diagnostics.</summary>
    public static async Task<PdfCore.PdfDocumentConversionResult> ToPdfDocumentResultAsync(this HtmlConversionDocument document, HtmlPdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        HtmlPdfRenderResult rendered = await HtmlPdfRenderedConverter.ConvertAsync(
            document.CreateDocumentForConversion(HtmlCssMediaContext.Print),
            Normalize(options),
            cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return CreateResult(rendered);
    }

    /// <summary>Asynchronously reads UTF-8 HTML from a stream and converts it to a PDF document plus diagnostics.</summary>
    public static async Task<PdfCore.PdfDocumentConversionResult> ToPdfDocumentResultAsync(this Stream htmlStream, HtmlPdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        if (htmlStream == null) throw new ArgumentNullException(nameof(htmlStream));
        string html = await HtmlTextIO.ReadAsync(htmlStream, cancellationToken).ConfigureAwait(false);
        return await html.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false);
    }

    private static HtmlPdfSaveOptions Normalize(HtmlPdfSaveOptions? options) => options?.ClonePdf() ?? new HtmlPdfSaveOptions();

    /// <summary>Serializes a completed conversion while honoring cancellation immediately before and after the synchronous writer.</summary>
    internal static byte[] SerializeToBytes(PdfCore.PdfDocumentConversionResult result, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        byte[] bytes = result.ToBytes();
        cancellationToken.ThrowIfCancellationRequested();
        return bytes;
    }

    private static PdfCore.PdfDocumentConversionResult CreateResult(HtmlPdfRenderResult rendered) {
        PdfCore.PdfConversionReport report = rendered.ConversionReport;
        foreach (HtmlDiagnostic diagnostic in rendered.Diagnostics.Diagnostics) {
            var details = string.IsNullOrWhiteSpace(diagnostic.Detail)
                ? null
                : new Dictionary<string, string> { ["Detail"] = diagnostic.Detail! };
            report.Add(new PdfCore.PdfConversionWarning(
                diagnostic.Component,
                diagnostic.Code,
                diagnostic.Source ?? "html-render",
                diagnostic.Message,
                MapSeverity(diagnostic.Severity),
                details: details));
        }
        return new PdfCore.PdfDocumentConversionResult(rendered.Document, report);
    }

    private static PdfCore.PdfConversionWarningSeverity MapSeverity(HtmlDiagnosticSeverity severity) => severity switch {
        HtmlDiagnosticSeverity.Info => PdfCore.PdfConversionWarningSeverity.Information,
        HtmlDiagnosticSeverity.Error => PdfCore.PdfConversionWarningSeverity.Error,
        _ => PdfCore.PdfConversionWarningSeverity.Warning
    };

}
