using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

/// <summary>Converts a parsed OfficeIMO HTML source document through the shared HTML render scene.</summary>
public static partial class HtmlPdfConverterExtensions {
    /// <summary>Converts a parsed HTML document to PDF plus an immutable diagnostics snapshot.</summary>
    internal static PdfCore.PdfDocumentConversionResult ConvertToPdfWithEvidence(this HtmlConversionDocument document, HtmlPdfExportOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentNullException.ThrowIfNull(document);
        HtmlPdfRenderResult rendered = HtmlPdfRenderedConverter.Convert(
            document,
            Normalize(options), cancellationToken);
        return CreateResult(rendered);
    }

    /// <summary>Asynchronously converts a parsed HTML document to PDF plus an immutable diagnostics snapshot.</summary>
    internal static async Task<PdfCore.PdfDocumentConversionResult> ConvertToPdfWithEvidenceAsync(
        this HtmlConversionDocument document,
        HtmlPdfExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(document);
        cancellationToken.ThrowIfCancellationRequested();
        HtmlPdfRenderResult rendered = await HtmlPdfRenderedConverter.ConvertAsync(
            document,
            Normalize(options),
            cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return CreateResult(rendered);
    }

    private static HtmlPdfExportOptions Normalize(HtmlPdfExportOptions? options) => options?.ClonePdf() ?? new HtmlPdfExportOptions();

    /// <summary>Serializes a completed conversion while honoring the supplied cancellation token.</summary>
    internal static byte[] SerializeToBytes(PdfCore.PdfDocumentConversionResult result, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        byte[] bytes = result.ToBytes(cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return bytes;
    }

    internal static PdfCore.PdfDocumentConversionResult CreateResult(HtmlPdfRenderResult rendered) {
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
                diagnostic.LossKind,
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
