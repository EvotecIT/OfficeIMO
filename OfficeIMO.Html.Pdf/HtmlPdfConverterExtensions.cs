using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

/// <summary>Converts a parsed OfficeIMO HTML source document through the shared HTML render scene.</summary>
public static partial class HtmlPdfConverterExtensions {
    /// <summary>Converts a parsed HTML document to PDF bytes.</summary>
    public static byte[] ToPdf(this HtmlConversionDocument document, HtmlPdfSaveOptions? options = null) =>
        document.ToPdfDocumentResult(options).ToBytes();

    /// <summary>Asynchronously resolves HTML resources and converts a parsed HTML document to PDF bytes.</summary>
    public static async Task<byte[]> ToPdfAsync(
        this HtmlConversionDocument document,
        HtmlPdfSaveOptions? options = null,
        CancellationToken cancellationToken = default) =>
        SerializeToBytes(await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false), cancellationToken);

    /// <summary>Converts a parsed HTML document to the first-party PDF document model.</summary>
    public static PdfCore.PdfDocument ToPdfDocument(this HtmlConversionDocument document, HtmlPdfSaveOptions? options = null) =>
        document.ToPdfDocumentResult(options).Value;

    /// <summary>Asynchronously converts a parsed HTML document to the first-party PDF document model.</summary>
    public static async Task<PdfCore.PdfDocument> ToPdfDocumentAsync(
        this HtmlConversionDocument document,
        HtmlPdfSaveOptions? options = null,
        CancellationToken cancellationToken = default) =>
        (await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false)).Value;

    /// <summary>Converts a parsed HTML document to PDF plus an immutable diagnostics snapshot.</summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this HtmlConversionDocument document, HtmlPdfSaveOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        HtmlPdfRenderResult rendered = HtmlPdfRenderedConverter.Convert(
            document,
            Normalize(options));
        return CreateResult(rendered);
    }

    /// <summary>Asynchronously converts a parsed HTML document to PDF plus an immutable diagnostics snapshot.</summary>
    public static async Task<PdfCore.PdfDocumentConversionResult> ToPdfDocumentResultAsync(
        this HtmlConversionDocument document,
        HtmlPdfSaveOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        HtmlPdfRenderResult rendered = await HtmlPdfRenderedConverter.ConvertAsync(
            document,
            Normalize(options),
            cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return CreateResult(rendered);
    }

    private static HtmlPdfSaveOptions Normalize(HtmlPdfSaveOptions? options) => options?.ClonePdf() ?? new HtmlPdfSaveOptions();

    /// <summary>Serializes a completed conversion while honoring cancellation around the synchronous writer.</summary>
    internal static byte[] SerializeToBytes(PdfCore.PdfDocumentConversionResult result, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        byte[] bytes = result.ToBytes();
        cancellationToken.ThrowIfCancellationRequested();
        return bytes;
    }

    private static PdfCore.PdfDocumentConversionResult CreateResult(HtmlPdfRenderResult rendered) {
        PdfCore.PdfConversionReport report = rendered.ConversionReport;
        foreach (HtmlDiagnostic diagnostic in rendered.Diagnostics) {
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
