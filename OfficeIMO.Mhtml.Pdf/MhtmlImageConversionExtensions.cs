using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Html.Pdf;

namespace OfficeIMO.Mhtml;

/// <summary>Renders MHTML through the existing PDF layout and shared image-export owners.</summary>
public static class MhtmlImageConversionExtensions {
    /// <summary>Renders selected pages, retaining MHTML conversion and PDF rendering diagnostics.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ToImages(this MhtmlDocument document,
        PdfImageExportOptions? imageOptions = null, HtmlPdfExportOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) =>
        CreateOperation(document, imageOptions, conversionOptions).Export(cancellationToken);

    /// <summary>Renders selected pages with an explicit encoding.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ToImages(this MhtmlDocument document, OfficeImageExportFormat format,
        PdfImageExportOptions? imageOptions = null, HtmlPdfExportOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) {
        PdfImageExportOptions options = imageOptions?.Clone() ?? new PdfImageExportOptions();
        options.Format = format;
        return document.ToImages(options, conversionOptions, cancellationToken);
    }

    /// <summary>Streams selected pages without retaining earlier encoded payloads.</summary>
    public static void ToImages(this MhtmlDocument document, OfficeImageExportConsumer consumer,
        PdfImageExportOptions? imageOptions = null, HtmlPdfExportOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) =>
        CreateOperation(document, imageOptions, conversionOptions).ExportEach(consumer, cancellationToken);

    /// <summary>Renders selected pages asynchronously under the shared deadline and cancellation policy.</summary>
    public static Task<IReadOnlyList<OfficeImageExportResult>> ToImagesAsync(this MhtmlDocument document,
        PdfImageExportOptions? imageOptions = null, HtmlPdfExportOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) =>
        CreateOperation(document, imageOptions, conversionOptions).ExportAsync(cancellationToken);

    /// <summary>Saves selected pages using shared atomic image files, limits, and loss rejection.</summary>
    public static OfficeConversionReport SaveAsImages(this MhtmlDocument document, string folderPath,
        PdfImageExportOptions? imageOptions = null, HtmlPdfExportOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        PdfImageExportOptions options = imageOptions?.Clone() ?? new PdfImageExportOptions();
        var saved = CreateOperation(document, options, conversionOptions).SaveFiles(folderPath, cancellationToken);
        return OfficeImageExportEvidence.CreateReport("MHTML", options.Format, saved.Report.NativeDiagnostics,
            saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    /// <summary>Saves selected pages asynchronously while releasing completed image payloads.</summary>
    public static async Task<OfficeConversionReport> SaveAsImagesAsync(this MhtmlDocument document, string folderPath,
        PdfImageExportOptions? imageOptions = null, HtmlPdfExportOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        PdfImageExportOptions options = imageOptions?.Clone() ?? new PdfImageExportOptions();
        var saved = await CreateOperation(document, options, conversionOptions).SaveFilesAsync(folderPath, cancellationToken).ConfigureAwait(false);
        return OfficeImageExportEvidence.CreateReport("MHTML", options.Format, saved.Report.NativeDiagnostics,
            saved.Report.ResultCount, timer.Elapsed, Path.GetFullPath(folderPath));
    }

    private static OfficeImageExportBatchOperation<PdfImageExportOptions> CreateOperation(MhtmlDocument document,
        PdfImageExportOptions? imageOptions, HtmlPdfExportOptions? conversionOptions) {
        ArgumentNullException.ThrowIfNull(document);
        PdfImageExportOptions image = imageOptions?.Clone() ?? new PdfImageExportOptions();
        HtmlPdfExportOptions conversion = conversionOptions?.ClonePdf() ?? new HtmlPdfExportOptions();
        image.FailOnLoss |= conversion.FailOnLoss;
        conversion.UseRenderingProfile(new OfficeRenderingProfile("mhtml-image", image.Fonts,
            image.TextShapingProvider, image.TextShapingLanguage, image.ImageCodec, image.Policy), OfficeRenderingProfileApplyMode.Overlay);
        return new OfficeImageExportBatchOperation<PdfImageExportOptions>(image, image.Format, image.FileConflictPolicy,
            (_, effective) => Collect(document, effective, conversion),
            exportEach: (_, effective, consumer, token) => Render(document, effective, conversion, consumer, token));
    }

    private static IReadOnlyList<OfficeImageExportResult> Collect(MhtmlDocument document, PdfImageExportOptions image,
        HtmlPdfExportOptions conversion) {
        var results = new List<OfficeImageExportResult>();
        Render(document, image, conversion, results.Add, default);
        return results.AsReadOnly();
    }

    private static void Render(MhtmlDocument document, PdfImageExportOptions image, HtmlPdfExportOptions conversion,
        OfficeImageExportConsumer consumer, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfDocumentConversionResult converted = document.ConvertToPdfWithEvidence(conversion, cancellationToken);
        using (converted.Value) {
            converted.Value.CreateImageExportOperation(image).ExportEach(result => {
                var diagnostics = PdfImageExportEngine.MapConversionDiagnostics(converted).Concat(result.Diagnostics).ToArray();
                consumer(image.EnsureAccepted(result.WithDiagnostics(diagnostics)));
            }, cancellationToken);
        }
    }
}
