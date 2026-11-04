using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Text;
using OfficeIMO.Markdown;
using OfficeIMO.Rtf;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Rtf.Markdown;

/// <summary>Semantic conversion between OfficeIMO RTF and Markdown document models.</summary>
public static partial class RtfMarkdownConverterExtensions {
    internal static RtfConversionResult<MarkdownDocument> ConvertToMarkdownWithEvidence(this RtfDocument document, RtfMarkdownExportOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        var context = new RtfToMarkdownConversionContext(options ?? new RtfMarkdownExportOptions());
        context.ConversionReport.Merge(document.GetStyleConversionDiagnostics());
        MarkdownDocument value = RtfToMarkdownConverter.Convert(document, context);
        return new RtfConversionResult<MarkdownDocument>(value, context.ConversionReport);
    }

    internal static RtfConversionResult<string> ConvertToMarkdownTextWithEvidence(this RtfDocument document, RtfMarkdownExportOptions? options = null) {
        RtfConversionResult<MarkdownDocument> converted = document.ConvertToMarkdownWithEvidence(options);
        return new RtfConversionResult<string>(converted.Value.ToMarkdown(options?.MarkdownWriteOptions), converted.Report);
    }

    internal static RtfConversionResult<RtfDocument> ConvertToRtfWithEvidence(this MarkdownDocument markdown, MarkdownRtfExportOptions? options = null) {
        if (markdown == null) throw new ArgumentNullException(nameof(markdown));
        var context = new MarkdownToRtfConversionContext(options ?? new MarkdownRtfExportOptions());
        RtfDocument value = MarkdownToRtfConverter.Convert(markdown, context);
        return new RtfConversionResult<RtfDocument>(value, context.ConversionReport);
    }

    /// <summary>Converts an RTF document into a Markdown document model.</summary>
    public static MarkdownDocument ToMarkdown(this RtfDocument document, RtfMarkdownExportOptions? options = null, CancellationToken cancellationToken = default) =>
        document.ToMarkdown(options, out _, cancellationToken);

    /// <summary>Converts an RTF document into a Markdown document model and returns the completed conversion report.</summary>
    public static MarkdownDocument ToMarkdown(this RtfDocument document, RtfMarkdownExportOptions? options, out OfficeConversionReport report,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        RtfConversionResult<MarkdownDocument> result = document.ConvertToMarkdownWithEvidence(options);
        MarkdownDocument value = result.Value;
        OfficeConversionReport source = result.Report;
        return OfficeConversionOutput.Complete(value, source, options?.FailOnLoss == true, out report);
    }

    /// <summary>Converts an RTF document into a Markdown document model, returning the value and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static Task<OfficeConversionResult<MarkdownDocument>> ToMarkdownAsync(this RtfDocument document, RtfMarkdownExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        return OfficeConversionOutput.RunAsync(() => (document.ToMarkdown(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);
    }

    /// <summary>Writes the converted output to a path and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsMarkdown(this RtfDocument document, string path, RtfMarkdownExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        MarkdownDocument value = document.ToMarkdown(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, path, OfficeConversionOutput.Text(value.ToMarkdown(options?.MarkdownWriteOptions), options?.Encoding));
    }

    /// <summary>Writes the converted output to a caller-owned stream, which stays open and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsMarkdown(this RtfDocument document, Stream stream, RtfMarkdownExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        MarkdownDocument value = document.ToMarkdown(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, stream, OfficeConversionOutput.Text(value.ToMarkdown(options?.MarkdownWriteOptions), options?.Encoding));
    }

    /// <summary>Asynchronously writes the converted output to a path and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsMarkdownAsync(this RtfDocument document, string path, RtfMarkdownExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            MarkdownDocument value = document.ToMarkdown(options, out OfficeConversionReport converted, cancellationToken);
            return (OfficeConversionOutput.Text(value.ToMarkdown(options?.MarkdownWriteOptions), options?.Encoding), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, path, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes the converted output to a caller-owned stream and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsMarkdownAsync(this RtfDocument document, Stream stream, RtfMarkdownExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            MarkdownDocument value = document.ToMarkdown(options, out OfficeConversionReport converted, cancellationToken);
            return (OfficeConversionOutput.Text(value.ToMarkdown(options?.MarkdownWriteOptions), options?.Encoding), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, stream, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Converts a Markdown document into an RTF document model.</summary>
    public static RtfDocument ToRtf(this MarkdownDocument markdown, MarkdownRtfExportOptions? options = null, CancellationToken cancellationToken = default) =>
        markdown.ToRtf(options, out _, cancellationToken);

    /// <summary>Converts a Markdown document into an RTF document model and returns the completed conversion report.</summary>
    public static RtfDocument ToRtf(this MarkdownDocument markdown, MarkdownRtfExportOptions? options, out OfficeConversionReport report,
        CancellationToken cancellationToken = default) {
        if (markdown == null) throw new ArgumentNullException(nameof(markdown));
        cancellationToken.ThrowIfCancellationRequested();
        RtfConversionResult<RtfDocument> result = markdown.ConvertToRtfWithEvidence(options);
        RtfDocument value = result.Value;
        OfficeConversionReport source = result.Report;
        return OfficeConversionOutput.Complete(value, source, options?.FailOnLoss == true, out report);
    }

    /// <summary>Converts a Markdown document into an RTF document model, returning the value and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static Task<OfficeConversionResult<RtfDocument>> ToRtfAsync(this MarkdownDocument markdown, MarkdownRtfExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        return OfficeConversionOutput.RunAsync(() => (markdown.ToRtf(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);
    }

    /// <summary>Writes the converted output to a path and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsRtf(this MarkdownDocument markdown, string path, MarkdownRtfExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        RtfDocument value = markdown.ToRtf(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, path, value.ToBytes(options?.SaveOptions, options?.Encoding));
    }

    /// <summary>Writes the converted output to a caller-owned stream, which stays open and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsRtf(this MarkdownDocument markdown, Stream stream, MarkdownRtfExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        RtfDocument value = markdown.ToRtf(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, stream, value.ToBytes(options?.SaveOptions, options?.Encoding));
    }

    /// <summary>Asynchronously writes the converted output to a path and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsRtfAsync(this MarkdownDocument markdown, string path, MarkdownRtfExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            RtfDocument value = markdown.ToRtf(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(options?.SaveOptions, options?.Encoding), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, path, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes the converted output to a caller-owned stream and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsRtfAsync(this MarkdownDocument markdown, Stream stream, MarkdownRtfExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            RtfDocument value = markdown.ToRtf(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(options?.SaveOptions, options?.Encoding), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, stream, bytes, cancellationToken).ConfigureAwait(false);
    }
}
