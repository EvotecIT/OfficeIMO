using OfficeIMO.Core.Internal;
using OfficeIMO.Html;
using System.Threading;
using System.Threading.Tasks;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Html;

public static partial class PowerPointHtmlConverterExtensions {
    /// <summary>Saves a presentation as UTF-8 HTML and returns its completed-stage report.</summary>
    public static OfficeConversionReport SaveAsHtml(this PptCore.PowerPointDocument presentation, string path,
        PowerPointHtmlSaveOptions? options = null, CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("HTML path cannot be empty.", nameof(path));
        string value = presentation.ToHtml(options, out OfficeConversionReport report, cancellationToken);
        OfficeConversionFailures.Run(report, () => HtmlTextIO.Write(path, value));
        return new OfficeConversionReport(report.Source, report.Target, path, report.ItemsWritten, report.Elapsed, report.Diagnostics);
    }

    /// <summary>Writes a presentation as UTF-8 HTML to a stream, leaves it open and returns its completed-stage report.</summary>
    public static OfficeConversionReport SaveAsHtml(this PptCore.PowerPointDocument presentation, Stream stream,
        PowerPointHtmlSaveOptions? options = null, CancellationToken cancellationToken = default) {
        string value = presentation.ToHtml(options, out OfficeConversionReport report, cancellationToken);
        OfficeConversionFailures.Run(report, () => HtmlTextIO.Write(stream, value));
        return new OfficeConversionReport(report.Source, report.Target, null, report.ItemsWritten, report.Elapsed, report.Diagnostics);
    }

    /// <summary>Asynchronously saves a presentation as UTF-8 HTML and returns its completed-stage report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the file is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsHtmlAsync(this PptCore.PowerPointDocument presentation, string path,
        PowerPointHtmlSaveOptions? options = null, CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("HTML path cannot be empty.", nameof(path));
        OfficeConversionResult<string> converted = await presentation.ToHtmlAsync(options, cancellationToken).ConfigureAwait(false);
        OfficeConversionReport report = converted.Report;
        await OfficeConversionFailures.RunAsync(report, () => HtmlTextIO.WriteAsync(path, converted.Value, cancellationToken)).ConfigureAwait(false);
        return new OfficeConversionReport(report.Source, report.Target, path, report.ItemsWritten, report.Elapsed, report.Diagnostics);
    }

    /// <summary>Asynchronously writes a presentation as UTF-8 HTML to a stream, leaves it open and returns its completed-stage report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the stream is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsHtmlAsync(this PptCore.PowerPointDocument presentation, Stream stream,
        PowerPointHtmlSaveOptions? options = null, CancellationToken cancellationToken = default) {
        OfficeConversionResult<string> converted = await presentation.ToHtmlAsync(options, cancellationToken).ConfigureAwait(false);
        OfficeConversionReport report = converted.Report;
        await OfficeConversionFailures.RunAsync(report, () => HtmlTextIO.WriteAsync(stream, converted.Value, cancellationToken)).ConfigureAwait(false);
        return new OfficeConversionReport(report.Source, report.Target, null, report.ItemsWritten, report.Elapsed, report.Diagnostics);
    }
}
