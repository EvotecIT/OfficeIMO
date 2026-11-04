using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Text;
using OfficeIMO.Core.Internal;
using OfficeIMO.Rtf;

namespace OfficeIMO.Html;

public static partial class HtmlRtfConverterExtensions {
    /// <summary>Converts an RTF document to semantic HTML text.</summary>
    public static string ToHtml(this RtfDocument document, RtfHtmlExportOptions? options = null, CancellationToken cancellationToken = default) =>
        document.ToHtml(options, out _, cancellationToken);

    /// <summary>Converts an RTF document to semantic HTML text and returns the completed conversion report.</summary>
    public static string ToHtml(this RtfDocument document, RtfHtmlExportOptions? options, out OfficeConversionReport report,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        RtfToHtmlResult result = document.ConvertToHtmlWithEvidence(options);
        string value = result.RequireValue();
        OfficeConversionReport source = result.Report;
        return OfficeConversionOutput.Complete(value, source, options?.FailOnLoss == true, out report);
    }

    /// <summary>Converts an RTF document to semantic HTML text, returning the value and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static Task<OfficeConversionResult<string>> ToHtmlAsync(this RtfDocument document, RtfHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        return OfficeConversionOutput.RunAsync(() => (document.ToHtml(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);
    }

    /// <summary>Writes the converted output to a path and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsHtml(this RtfDocument document, string path, RtfHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        string value = document.ToHtml(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, path, OfficeConversionOutput.Text(value, options?.Encoding));
    }

    /// <summary>Writes the converted output to a caller-owned stream, which stays open and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsHtml(this RtfDocument document, Stream stream, RtfHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        string value = document.ToHtml(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, stream, OfficeConversionOutput.Text(value, options?.Encoding));
    }

    /// <summary>Asynchronously writes the converted output to a path and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsHtmlAsync(this RtfDocument document, string path, RtfHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            string value = document.ToHtml(options, out OfficeConversionReport converted, cancellationToken);
            return (OfficeConversionOutput.Text(value, options?.Encoding), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, path, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes the converted output to a caller-owned stream and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsHtmlAsync(this RtfDocument document, Stream stream, RtfHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            string value = document.ToHtml(options, out OfficeConversionReport converted, cancellationToken);
            return (OfficeConversionOutput.Text(value, options?.Encoding), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, stream, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Converts prepared HTML into an RTF document model.</summary>
    public static RtfDocument ToRtf(this HtmlConversionDocument document, HtmlRtfExportOptions? options = null, CancellationToken cancellationToken = default) =>
        document.ToRtf(options, out _, cancellationToken);

    /// <summary>Converts prepared HTML into an RTF document model and returns the completed conversion report.</summary>
    public static RtfDocument ToRtf(this HtmlConversionDocument document, HtmlRtfExportOptions? options, out OfficeConversionReport report,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        HtmlToRtfResult result = document.ConvertToRtfWithEvidence(options);
        RtfDocument value = result.RequireValue();
        OfficeConversionReport source = result.Report;
        return OfficeConversionOutput.Complete(value, source, options?.FailOnLoss == true, out report);
    }

    /// <summary>Converts prepared HTML into an RTF document model, returning the value and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static Task<OfficeConversionResult<RtfDocument>> ToRtfAsync(this HtmlConversionDocument document, HtmlRtfExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        return OfficeConversionOutput.RunAsync(() => (document.ToRtf(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);
    }

    /// <summary>Writes the converted output to a path and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsRtf(this HtmlConversionDocument document, string path, HtmlRtfExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        RtfDocument value = document.ToRtf(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, path, value.ToBytes(options?.SaveOptions, options?.Encoding));
    }

    /// <summary>Writes the converted output to a caller-owned stream, which stays open and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsRtf(this HtmlConversionDocument document, Stream stream, HtmlRtfExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        RtfDocument value = document.ToRtf(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, stream, value.ToBytes(options?.SaveOptions, options?.Encoding));
    }

    /// <summary>Asynchronously writes the converted output to a path and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsRtfAsync(this HtmlConversionDocument document, string path, HtmlRtfExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            RtfDocument value = document.ToRtf(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(options?.SaveOptions, options?.Encoding), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, path, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes the converted output to a caller-owned stream and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsRtfAsync(this HtmlConversionDocument document, Stream stream, HtmlRtfExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            RtfDocument value = document.ToRtf(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(options?.SaveOptions, options?.Encoding), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, stream, bytes, cancellationToken).ConfigureAwait(false);
    }
}
