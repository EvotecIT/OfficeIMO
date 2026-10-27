using System.Diagnostics;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Core.Internal;

namespace OfficeIMO;

public static partial class OfficeOcrConversionExtensions {
    /// <summary>Renders a PDF as HTML.</summary>
    public static string ToHtml(this PdfOcrMergeResult source, PdfToHtmlOptions? options = null,
        CancellationToken cancellationToken = default) => source.ToHtml(options, out _, cancellationToken);

    /// <summary>Renders a PDF as HTML and returns its completed fidelity findings.</summary>
    public static string ToHtml(this PdfOcrMergeResult source, PdfToHtmlOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        PdfHtmlConversionResult result = ReadOcrHtml(source, options, cancellationToken);
        report = CompleteHtml(result, options, timer.Elapsed);
        return result.Value;
    }

    /// <summary>Returns HTML and its completed fidelity findings together.</summary>
    public static async Task<OfficeConversionResult<string>> ToHtmlAsync(this PdfOcrMergeResult source,
        PdfToHtmlOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new PdfToHtmlOptions()).CloneForConversion();
        return await Task.Run(() => {
            string value = source.ToHtml(options, out OfficeConversionReport report, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            return new OfficeConversionResult<string>(value, report);
        }, cancellationToken).ConfigureAwait(false);
    }

    private static PdfHtmlConversionResult ReadOcrHtml(PdfOcrMergeResult source, PdfToHtmlOptions? options, CancellationToken token) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        token.ThrowIfCancellationRequested();
        source.RequireAcceptedOcrContent();
        return source.Document.ToHtmlResult(options, token);
    }

    private static OfficeConversionReport CompleteHtml(PdfHtmlConversionResult result, PdfToHtmlOptions? options,
        TimeSpan elapsed, string? path = null) {
        OfficeConversionReport report = PdfConversionEvidence.CreateReport("Pdf", new IOfficeConversionReport[] { result.Report },
            path == null ? null : Path.GetFullPath(path), elapsed, result.Summary.RenderedPageCount, "Html");
        if (options?.FailOnLoss == true) report.ThrowIfLoss();
        return report;
    }

    /// <summary>Writes UTF-8 HTML to a file and returns its completed conversion findings.</summary>
    public static OfficeConversionReport SaveAsHtml(this PdfOcrMergeResult source, string path,
        PdfToHtmlOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new PdfToHtmlOptions()).CloneForConversion();
        var timer = Stopwatch.StartNew();
        PdfHtmlConversionResult result = ReadOcrHtml(source, options, cancellationToken);
        OfficeConversionReport report = CompleteHtml(result, options, timer.Elapsed, path);
        cancellationToken.ThrowIfCancellationRequested();
        global::OfficeIMO.Core.Internal.OfficeConversionFailures.Run(report, () => OfficeFileCommit.WriteAllBytes(path, new UTF8Encoding(false).GetBytes(result.Value)));
        return new OfficeConversionReport("Pdf", "Html", report.OutputPath, report.ItemsWritten, timer.Elapsed, report.Diagnostics);
    }

    /// <summary>Writes UTF-8 HTML to a file and returns its completed conversion findings.</summary>
    public static async Task<OfficeConversionReport> SaveAsHtmlAsync(this PdfOcrMergeResult source, string path,
        PdfToHtmlOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new PdfToHtmlOptions()).CloneForConversion();
        var timer = Stopwatch.StartNew();
        PdfHtmlConversionResult result = await Task.Run(() => ReadOcrHtml(source, options, cancellationToken), cancellationToken).ConfigureAwait(false);
        OfficeConversionReport report = CompleteHtml(result, options, timer.Elapsed, path);
        cancellationToken.ThrowIfCancellationRequested();
        await global::OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(report, () => OfficeFileCommit.WriteAllBytesAsync(path, new UTF8Encoding(false).GetBytes(result.Value), cancellationToken: cancellationToken)).ConfigureAwait(false);
        return new OfficeConversionReport("Pdf", "Html", report.OutputPath, report.ItemsWritten, timer.Elapsed, report.Diagnostics);
    }

    /// <summary>Writes UTF-8 HTML to a caller-owned stream and returns its completed conversion findings.</summary>
    public static OfficeConversionReport SaveAsHtml(this PdfOcrMergeResult source, Stream stream,
        PdfToHtmlOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new PdfToHtmlOptions()).CloneForConversion();
        var timer = Stopwatch.StartNew();
        PdfHtmlConversionResult result = ReadOcrHtml(source, options, cancellationToken);
        OfficeConversionReport report = CompleteHtml(result, options, timer.Elapsed, null);
        cancellationToken.ThrowIfCancellationRequested();
        global::OfficeIMO.Core.Internal.OfficeConversionFailures.Run(report, () => OfficeStreamWriter.WriteAllBytes(stream, new UTF8Encoding(false).GetBytes(result.Value)));
        return new OfficeConversionReport("Pdf", "Html", report.OutputPath, report.ItemsWritten, timer.Elapsed, report.Diagnostics);
    }

    /// <summary>Writes UTF-8 HTML to a caller-owned stream and returns its completed conversion findings.</summary>
    public static async Task<OfficeConversionReport> SaveAsHtmlAsync(this PdfOcrMergeResult source, Stream stream,
        PdfToHtmlOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new PdfToHtmlOptions()).CloneForConversion();
        var timer = Stopwatch.StartNew();
        PdfHtmlConversionResult result = await Task.Run(() => ReadOcrHtml(source, options, cancellationToken), cancellationToken).ConfigureAwait(false);
        OfficeConversionReport report = CompleteHtml(result, options, timer.Elapsed, null);
        cancellationToken.ThrowIfCancellationRequested();
        await global::OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(report, () => OfficeStreamWriter.WriteAllBytesAsync(stream, new UTF8Encoding(false).GetBytes(result.Value), cancellationToken)).ConfigureAwait(false);
        return new OfficeConversionReport("Pdf", "Html", report.OutputPath, report.ItemsWritten, timer.Elapsed, report.Diagnostics);
    }
}
