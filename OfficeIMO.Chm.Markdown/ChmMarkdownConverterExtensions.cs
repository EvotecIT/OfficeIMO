using OfficeIMO.Core.Internal;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Markdown;

namespace OfficeIMO.Chm;

/// <summary>Converts compiled help through the canonical HTML-to-Markdown engine.</summary>
public static class ChmMarkdownConverterExtensions {
    /// <summary>Exports selected topics in book order as an editable Markdown document.</summary>
    public static MarkdownDocument ToMarkdown(this ChmDocument document, ChmMarkdownExportOptions? options = null,
        CancellationToken cancellationToken = default) => document.ToMarkdown(options, out _, cancellationToken);

    /// <summary>Exports selected topics and returns topic, projection and Markdown findings.</summary>
    public static MarkdownDocument ToMarkdown(this ChmDocument document, ChmMarkdownExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        var (value, _, source) = Convert(document, options, cancellationToken);
        return OfficeConversionOutput.Complete(value, source, options?.FailOnLoss == true, out report);
    }

    /// <summary>Exports selected topics with the report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static Task<OfficeConversionResult<MarkdownDocument>> ToMarkdownAsync(this ChmDocument document,
        ChmMarkdownExportOptions? options = null, CancellationToken cancellationToken = default) {
        options = options?.Clone();
        return OfficeConversionOutput.RunAsync(() => (document.ToMarkdown(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);
    }

    /// <summary>Writes UTF-8 Markdown to a path and returns the report.</summary>
    public static OfficeConversionReport SaveAsMarkdown(this ChmDocument document, string path, ChmMarkdownExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        var (_, text, report) = Write(document, options, cancellationToken);
        return OfficeConversionOutput.Save(report, path, OfficeConversionOutput.Text(text));
    }

    /// <summary>Writes UTF-8 Markdown to a caller-owned stream and returns the report.</summary>
    public static OfficeConversionReport SaveAsMarkdown(this ChmDocument document, Stream stream, ChmMarkdownExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        var (_, text, report) = Write(document, options, cancellationToken);
        return OfficeConversionOutput.Save(report, stream, OfficeConversionOutput.Text(text));
    }

    /// <summary>Asynchronously writes UTF-8 Markdown to a path.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the file is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsMarkdownAsync(this ChmDocument document, string path,
        ChmMarkdownExportOptions? options = null, CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (_, text, report) = await Task.Run(() => Write(document, options, cancellationToken), cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, path, OfficeConversionOutput.Text(text), cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes UTF-8 Markdown to a caller-owned stream.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the stream is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsMarkdownAsync(this ChmDocument document, Stream stream,
        ChmMarkdownExportOptions? options = null, CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (_, text, report) = await Task.Run(() => Write(document, options, cancellationToken), cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, stream, OfficeConversionOutput.Text(text), cancellationToken).ConfigureAwait(false);
    }

    internal static ChmConversionResult<string> ConvertToMarkdownWithEvidence(this ChmDocument document, ChmConversionOptions? options = null,
        HtmlMarkdownExportOptions? markdownOptions = null, CancellationToken cancellationToken = default) {
        var (_, text, report) = Convert(document, new ChmMarkdownExportOptions { ConversionOptions = options ?? new ChmConversionOptions(), MarkdownOptions = markdownOptions }, cancellationToken);
        return new ChmConversionResult<string>(text, report);
    }

    internal static ChmConversionResult<MarkdownDocument> ConvertToMarkdownDocumentWithEvidence(this ChmDocument document, ChmConversionOptions? options = null,
        HtmlMarkdownExportOptions? markdownOptions = null, CancellationToken cancellationToken = default) {
        var (value, _, report) = Convert(document, new ChmMarkdownExportOptions { ConversionOptions = options ?? new ChmConversionOptions(), MarkdownOptions = markdownOptions }, cancellationToken);
        return new ChmConversionResult<MarkdownDocument>(value, report);
    }

    private static (MarkdownDocument, string, OfficeConversionReport) Write(ChmDocument document, ChmMarkdownExportOptions? options,
        CancellationToken cancellationToken) {
        var (value, text, report) = Convert(document, options, cancellationToken);
        if (options?.FailOnLoss == true) report.ThrowIfLoss();
        return (value, text, report);
    }

    private static (MarkdownDocument Value, string Text, ChmConversionReport Report) Convert(ChmDocument document,
        ChmMarkdownExportOptions? options, CancellationToken cancellationToken) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        ChmConversionOptions configured = options?.ConversionOptions.Clone() ?? new ChmConversionOptions();
        ChmConversionResult<HtmlConversionDocument> projected = document.ToHtmlDocumentResult(configured, cancellationToken);
        HtmlMarkdownExportOptions conversion = options?.MarkdownOptions?.Clone() ?? HtmlMarkdownExportOptions.CreatePortableProfile();
        conversion.MaxInputCharacters = Math.Min(conversion.MaxInputCharacters ?? int.MaxValue, (int)configured.MaxOutputBytes);
        var converted = HtmlMarkdownConverterExtensions.ConvertToMarkdownWithEvidence(projected.Value, conversion, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        string markdown = converted.Value.ToMarkdown(conversion.MarkdownWriteOptions, cancellationToken);
        ChmDocument.EnforceOutput(Encoding.UTF8.GetByteCount(markdown), configured);
        cancellationToken.ThrowIfCancellationRequested();
        var report = new ChmConversionReport(projected.Report.TopicPaths,
            projected.Report.Diagnostics.Concat(converted.Report.Diagnostics));
        return (converted.Value, markdown, report);
    }
}
