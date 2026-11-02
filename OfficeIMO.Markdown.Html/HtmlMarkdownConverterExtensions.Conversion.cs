using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Html;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Markdown.Html;

public static partial class HtmlMarkdownConverterExtensions {
    /// <summary>Converts prepared HTML to a Markdown document.</summary>
    public static MarkdownDocument ToMarkdown(this HtmlConversionDocument document, MarkdownHtmlImportOptions? options = null, CancellationToken cancellationToken = default) =>
        document.ToMarkdown(options, out _, cancellationToken);
    /// <summary>Converts prepared HTML to Markdown with completed-stage findings.</summary>
    public static MarkdownDocument ToMarkdown(this HtmlConversionDocument document, MarkdownHtmlImportOptions? options, out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        var result = ConvertToMarkdownWithEvidence(document, options?.Clone() ?? new MarkdownHtmlImportOptions(), cancellationToken);
        report = HtmlConversionEvidence.CreateReport("Html", "Markdown", result.Report.Diagnostics, timer.Elapsed);
        if (options?.FailOnLoss == true && report.HasLoss) { result.Value.Dispose(); report.ThrowIfLoss(); }
        return result.RequireValue();
    }
    /// <summary>Converts HTML to Markdown and returns the model and report together.</summary>
    public static async Task<OfficeConversionResult<MarkdownDocument>> ToMarkdownAsync(this HtmlConversionDocument document, MarkdownHtmlImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new MarkdownHtmlImportOptions()).Clone();
        return await Task.Run(() => { var value = document.ToMarkdown(options, out var report, cancellationToken); try { cancellationToken.ThrowIfCancellationRequested(); }
            catch { value.Dispose(); throw; }
            return new OfficeConversionResult<MarkdownDocument>(value, report); }, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Saves converted Markdown with completed-stage findings.</summary>
    public static OfficeConversionReport SaveAsMarkdown(this HtmlConversionDocument document, string path, MarkdownHtmlImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = options?.Clone() ?? new MarkdownHtmlImportOptions();
        using var value = document.ToMarkdown(options, out var report, cancellationToken);
        var saveOptions = options.MarkdownWriteOptions?.Clone() ?? new MarkdownSaveOptions();
        saveOptions.Encoding = options.Encoding;
        saveOptions.AllowLoss = !options.FailOnLoss;
        OfficeConversionReport output = value.AssessNativeSave(saveOptions);
        if (options.FailOnLoss) output.ThrowIfLoss();
        report = new OfficeConversionReport(report.Source, report.Target, itemsWritten: report.ItemsWritten,
            elapsed: report.Elapsed, diagnostics: report.Diagnostics.Concat(output.Diagnostics));
        byte[] bytes = value.ToBytes(saveOptions, cancellationToken);
        global::OfficeIMO.Core.Internal.OfficeConversionFailures.Run(report, () => OfficeFileCommit.Write(path, output => OfficeStreamWriter.CopyBytes(output, bytes, cancellationToken), cancellationToken));
        return new OfficeConversionReport(report.Source, report.Target, path, report.ItemsWritten, report.Elapsed, report.Diagnostics);
    }

    /// <summary>Saves converted Markdown with completed-stage findings.</summary>
    public static async Task<OfficeConversionReport> SaveAsMarkdownAsync(this HtmlConversionDocument document, string path, MarkdownHtmlImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = options?.Clone() ?? new MarkdownHtmlImportOptions();
        using var value = document.ToMarkdown(options, out var report, cancellationToken);
        var saveOptions = options.MarkdownWriteOptions?.Clone() ?? new MarkdownSaveOptions();
        saveOptions.Encoding = options.Encoding;
        saveOptions.AllowLoss = !options.FailOnLoss;
        OfficeConversionReport output = value.AssessNativeSave(saveOptions);
        if (options.FailOnLoss) output.ThrowIfLoss();
        report = new OfficeConversionReport(report.Source, report.Target, itemsWritten: report.ItemsWritten,
            elapsed: report.Elapsed, diagnostics: report.Diagnostics.Concat(output.Diagnostics));
        byte[] bytes = value.ToBytes(saveOptions, cancellationToken);
        await global::OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(report, () => OfficeFileCommit.WriteAllBytesAsync(path, bytes, cancellationToken: cancellationToken)).ConfigureAwait(false);
        return new OfficeConversionReport(report.Source, report.Target, path, report.ItemsWritten, report.Elapsed, report.Diagnostics);
    }

    /// <summary>Saves converted Markdown with completed-stage findings.</summary>
    public static OfficeConversionReport SaveAsMarkdown(this HtmlConversionDocument document, Stream stream, MarkdownHtmlImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = options?.Clone() ?? new MarkdownHtmlImportOptions();
        using var value = document.ToMarkdown(options, out var report, cancellationToken);
        var saveOptions = options.MarkdownWriteOptions?.Clone() ?? new MarkdownSaveOptions();
        saveOptions.Encoding = options.Encoding;
        saveOptions.AllowLoss = !options.FailOnLoss;
        OfficeConversionReport output = value.AssessNativeSave(saveOptions);
        if (options.FailOnLoss) output.ThrowIfLoss();
        report = new OfficeConversionReport(report.Source, report.Target, itemsWritten: report.ItemsWritten,
            elapsed: report.Elapsed, diagnostics: report.Diagnostics.Concat(output.Diagnostics));
        byte[] bytes = value.ToBytes(saveOptions, cancellationToken);
        global::OfficeIMO.Core.Internal.OfficeConversionFailures.Run(report, () => OfficeStreamWriter.Write(stream, output => OfficeStreamWriter.CopyBytes(output, bytes, cancellationToken)));
        return new OfficeConversionReport(report.Source, report.Target, null, report.ItemsWritten, report.Elapsed, report.Diagnostics);
    }

    /// <summary>Saves converted Markdown with completed-stage findings.</summary>
    public static async Task<OfficeConversionReport> SaveAsMarkdownAsync(this HtmlConversionDocument document, Stream stream, MarkdownHtmlImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = options?.Clone() ?? new MarkdownHtmlImportOptions();
        using var value = document.ToMarkdown(options, out var report, cancellationToken);
        var saveOptions = options.MarkdownWriteOptions?.Clone() ?? new MarkdownSaveOptions();
        saveOptions.Encoding = options.Encoding;
        saveOptions.AllowLoss = !options.FailOnLoss;
        OfficeConversionReport output = value.AssessNativeSave(saveOptions);
        if (options.FailOnLoss) output.ThrowIfLoss();
        report = new OfficeConversionReport(report.Source, report.Target, itemsWritten: report.ItemsWritten,
            elapsed: report.Elapsed, diagnostics: report.Diagnostics.Concat(output.Diagnostics));
        byte[] bytes = value.ToBytes(saveOptions, cancellationToken);
        await global::OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(report, () => OfficeStreamWriter.WriteAllBytesAsync(stream, bytes, cancellationToken: cancellationToken)).ConfigureAwait(false);
        return new OfficeConversionReport(report.Source, report.Target, null, report.ItemsWritten, report.Elapsed, report.Diagnostics);
    }
}
