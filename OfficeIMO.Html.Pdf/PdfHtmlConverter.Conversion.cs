using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

public static partial class PdfHtmlConverterExtensions {
    /// <summary>Renders a PDF as HTML.</summary>
    public static string ToHtml(this PdfDocument document, PdfToHtmlOptions? options = null,
        CancellationToken cancellationToken = default) => document.ToHtml(options, out _, cancellationToken);

    /// <summary>Renders a PDF as HTML and returns its completed fidelity findings.</summary>
    public static string ToHtml(this PdfDocument document, PdfToHtmlOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        PdfHtmlConversionResult result = document.ToHtmlResult(options, cancellationToken);
        report = CompleteHtml(result, options, timer.Elapsed);
        return result.Value;
    }

    /// <summary>Returns HTML and its completed fidelity findings together.</summary>
    public static async Task<OfficeConversionResult<string>> ToHtmlAsync(this PdfDocument document,
        PdfToHtmlOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new PdfToHtmlOptions()).CloneForConversion();
        return await Task.Run(() => {
            string value = document.ToHtml(options, out OfficeConversionReport report, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            return new OfficeConversionResult<string>(value, report);
        }, cancellationToken).ConfigureAwait(false);
    }

    private static OfficeConversionReport CompleteHtml(PdfHtmlConversionResult result, PdfToHtmlOptions? options,
        TimeSpan elapsed, string? path = null) {
        OfficeConversionReport report = PdfConversionEvidence.CreateReport("Pdf", new IOfficeConversionReport[] { result.Report },
            path == null ? null : Path.GetFullPath(path), elapsed, result.Summary.RenderedPageCount, "Html");
        if (options?.FailOnLoss == true) report.ThrowIfLoss();
        return report;
    }

    /// <summary>Writes UTF-8 HTML to a file and returns its completed conversion findings.</summary>
    public static OfficeConversionReport SaveAsHtml(this PdfDocument document, string path,
        PdfToHtmlOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new PdfToHtmlOptions()).CloneForConversion();
        var timer = Stopwatch.StartNew();
        PdfHtmlConversionResult result = document.ToHtmlResult(options, cancellationToken);
        OfficeConversionReport report = CompleteHtml(result, options, timer.Elapsed, path);
        cancellationToken.ThrowIfCancellationRequested();
        global::OfficeIMO.Core.Internal.OfficeConversionFailures.Run(report, () => HtmlTextIO.Write(path, result.Value));
        return new OfficeConversionReport("Pdf", "Html", report.OutputPath, report.ItemsWritten, timer.Elapsed, report.Diagnostics);
    }

    /// <summary>Writes UTF-8 HTML to a file and returns its completed conversion findings.</summary>
    public static async Task<OfficeConversionReport> SaveAsHtmlAsync(this PdfDocument document, string path,
        PdfToHtmlOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new PdfToHtmlOptions()).CloneForConversion();
        var timer = Stopwatch.StartNew();
        PdfHtmlConversionResult result = await Task.Run(() => document.ToHtmlResult(options, cancellationToken), cancellationToken).ConfigureAwait(false);
        OfficeConversionReport report = CompleteHtml(result, options, timer.Elapsed, path);
        cancellationToken.ThrowIfCancellationRequested();
        await global::OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(report, () => HtmlTextIO.WriteAsync(path, result.Value, cancellationToken)).ConfigureAwait(false);
        return new OfficeConversionReport("Pdf", "Html", report.OutputPath, report.ItemsWritten, timer.Elapsed, report.Diagnostics);
    }

    /// <summary>Writes UTF-8 HTML to a caller-owned stream and returns its completed conversion findings.</summary>
    public static OfficeConversionReport SaveAsHtml(this PdfDocument document, Stream stream,
        PdfToHtmlOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new PdfToHtmlOptions()).CloneForConversion();
        var timer = Stopwatch.StartNew();
        PdfHtmlConversionResult result = document.ToHtmlResult(options, cancellationToken);
        OfficeConversionReport report = CompleteHtml(result, options, timer.Elapsed, null);
        cancellationToken.ThrowIfCancellationRequested();
        global::OfficeIMO.Core.Internal.OfficeConversionFailures.Run(report, () => HtmlTextIO.Write(stream, result.Value));
        return new OfficeConversionReport("Pdf", "Html", report.OutputPath, report.ItemsWritten, timer.Elapsed, report.Diagnostics);
    }

    /// <summary>Writes UTF-8 HTML to a caller-owned stream and returns its completed conversion findings.</summary>
    public static async Task<OfficeConversionReport> SaveAsHtmlAsync(this PdfDocument document, Stream stream,
        PdfToHtmlOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new PdfToHtmlOptions()).CloneForConversion();
        var timer = Stopwatch.StartNew();
        PdfHtmlConversionResult result = await Task.Run(() => document.ToHtmlResult(options, cancellationToken), cancellationToken).ConfigureAwait(false);
        OfficeConversionReport report = CompleteHtml(result, options, timer.Elapsed, null);
        cancellationToken.ThrowIfCancellationRequested();
        await global::OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(report, () => HtmlTextIO.WriteAsync(stream, result.Value, cancellationToken)).ConfigureAwait(false);
        return new OfficeConversionReport("Pdf", "Html", report.OutputPath, report.ItemsWritten, timer.Elapsed, report.Diagnostics);
    }
}
