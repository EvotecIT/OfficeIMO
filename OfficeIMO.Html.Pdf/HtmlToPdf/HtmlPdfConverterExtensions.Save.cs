using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

public static partial class HtmlPdfConverterExtensions {
    /// <summary>Converts the source document to a PDF model.</summary>
    public static PdfDocument ToPdf(this HtmlConversionDocument document, HtmlToPdfOptions? options = null,
        CancellationToken cancellationToken = default) =>
        document.ToPdf(options, out _, cancellationToken);

    /// <summary>Converts the source document and returns its completed conversion findings.</summary>
    public static PdfDocument ToPdf(this HtmlConversionDocument document, HtmlToPdfOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        options = Normalize(options);
        var timer = Stopwatch.StartNew();
        OfficeConversionResult<PdfDocument> completed = HtmlPdfRenderedConverter.Complete(document, options,
            cancellationToken, (rendered, token) => {
                PdfDocument value = PdfConversionEvidence.CompleteModel(CreateResult(rendered), "Html", options.FailOnLoss,
                    timer, token, out OfficeConversionReport findings);
                return new OfficeConversionResult<PdfDocument>(value, findings);
            });
        report = completed.Report;
        return completed.Value;
    }

    /// <summary>Resolves resources asynchronously and returns the PDF model and completed conversion report.</summary>
    public static async Task<OfficeConversionResult<PdfDocument>> ToPdfAsync(this HtmlConversionDocument document,
        HtmlToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        options = Normalize(options);
        var timer = Stopwatch.StartNew();
        return await HtmlPdfRenderedConverter.CompleteAsync(document, options, cancellationToken, (rendered, token) => {
            PdfDocument value = PdfConversionEvidence.CompleteModel(CreateResult(rendered), "Html", options.FailOnLoss,
                timer, token, out OfficeConversionReport report);
            return Task.FromResult(new OfficeConversionResult<PdfDocument>(value, report));
        }).ConfigureAwait(false);
    }

    /// <summary>Writes a PDF file and returns the complete conversion report.</summary>
    public static OfficeConversionReport SaveAsPdf(this HtmlConversionDocument document, string path,
        HtmlToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(path);
        options = Normalize(options);
        var timer = Stopwatch.StartNew();
        return HtmlPdfRenderedConverter.Complete(document, options, cancellationToken, (rendered, token) => {
            PdfDocumentConversionResult converted = CreateResult(rendered);
            using PdfDocument model = converted.Value;
            PdfSaveResult saved = options.FailOnLoss
                ? PdfConversionEvidence.SaveLossless(converted, "Html", path, timer, token)
                : PdfConversionEvidence.RunOutput(converted, "Html", () => converted.Save(path, token));
            return PdfConversionEvidence.CreateReport("Html", saved.ConversionReports, saved.OutputPath, timer.Elapsed, saved.Pipeline.Output?.PageCount ?? 0);
        });
    }

    /// <summary>Writes PDF to a caller-owned stream and returns the complete conversion report.</summary>
    public static OfficeConversionReport SaveAsPdf(this HtmlConversionDocument document, Stream stream,
        HtmlToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(stream);
        options = Normalize(options);
        var timer = Stopwatch.StartNew();
        return HtmlPdfRenderedConverter.Complete(document, options, cancellationToken, (rendered, token) => {
            PdfDocumentConversionResult converted = CreateResult(rendered);
            using PdfDocument model = converted.Value;
            PdfSaveResult saved = options.FailOnLoss
                ? PdfConversionEvidence.SaveLossless(converted, "Html", stream, timer, token)
                : PdfConversionEvidence.RunOutput(converted, "Html", () => converted.Save(stream, token));
            return PdfConversionEvidence.CreateReport("Html", saved.ConversionReports, elapsed: timer.Elapsed, itemsWritten: saved.Pipeline.Output?.PageCount ?? 0);
        });
    }

    /// <summary>Asynchronously writes a PDF file and returns the complete conversion report.</summary>
    public static async Task<OfficeConversionReport> SaveAsPdfAsync(this HtmlConversionDocument document, string path,
        HtmlToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(path);
        options = Normalize(options);
        var timer = Stopwatch.StartNew();
        return await HtmlPdfRenderedConverter.CompleteAsync(document, options, cancellationToken, async (rendered, token) => {
            PdfDocumentConversionResult converted = CreateResult(rendered);
            using PdfDocument model = converted.Value;
            PdfSaveResult saved = options.FailOnLoss
                ? await PdfConversionEvidence.SaveLosslessAsync(converted, "Html", path, timer, token).ConfigureAwait(false)
                : await PdfConversionEvidence.RunOutputAsync(converted, "Html", () => converted.SaveAsync(path, token)).ConfigureAwait(false);
            return PdfConversionEvidence.CreateReport("Html", saved.ConversionReports, saved.OutputPath, timer.Elapsed, saved.Pipeline.Output?.PageCount ?? 0);
        }).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes PDF to a caller-owned stream and returns its conversion report.</summary>
    public static async Task<OfficeConversionReport> SaveAsPdfAsync(this HtmlConversionDocument document, Stream stream,
        HtmlToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(stream);
        options = Normalize(options);
        var timer = Stopwatch.StartNew();
        return await HtmlPdfRenderedConverter.CompleteAsync(document, options, cancellationToken, async (rendered, token) => {
            PdfDocumentConversionResult converted = CreateResult(rendered);
            using PdfDocument model = converted.Value;
            PdfSaveResult saved = options.FailOnLoss
                ? await PdfConversionEvidence.SaveLosslessAsync(converted, "Html", stream, timer, token).ConfigureAwait(false)
                : await PdfConversionEvidence.RunOutputAsync(converted, "Html", () => converted.SaveAsync(stream, token)).ConfigureAwait(false);
            return PdfConversionEvidence.CreateReport("Html", saved.ConversionReports, elapsed: timer.Elapsed, itemsWritten: saved.Pipeline.Output?.PageCount ?? 0);
        }).ConfigureAwait(false);
    }
}
