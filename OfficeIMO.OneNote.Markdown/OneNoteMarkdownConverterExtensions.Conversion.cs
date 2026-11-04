using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;
using OfficeIMO.Markdown;
using OfficeIMO.OneNote;

namespace OfficeIMO.OneNote.Markdown;

public static partial class OneNoteMarkdownConverterExtensions {
    /// <summary>Converts a section to a Markdown document.</summary>
    public static MarkdownDocument ToMarkdown(this OneNoteSection section, OneNoteMarkdownExportOptions? options = null, CancellationToken cancellationToken = default) =>
        section.ToMarkdown(options, out _, cancellationToken);

    /// <summary>Converts a section to a Markdown document and returns the completed conversion report.</summary>
    public static MarkdownDocument ToMarkdown(this OneNoteSection section, OneNoteMarkdownExportOptions? options, out OfficeConversionReport report,
        CancellationToken cancellationToken = default) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        cancellationToken.ThrowIfCancellationRequested();
        OneNoteMarkdownConversionResult result = ConvertToMarkdownWithEvidence(section, options, out _);
        MarkdownDocument value = result.Value;
        OfficeConversionReport source = result.Report;
        return OfficeConversionOutput.Complete(value, source, options?.FailOnLoss == true, out report);
    }

    /// <summary>Converts a section to a Markdown document, returning the value and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static Task<OfficeConversionResult<MarkdownDocument>> ToMarkdownAsync(this OneNoteSection section, OneNoteMarkdownExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        return OfficeConversionOutput.RunAsync(() => (section.ToMarkdown(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);
    }

    /// <summary>Writes the section's Markdown projection as UTF-8 to a path and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsMarkdown(this OneNoteSection section, string path, OneNoteMarkdownExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        var (bytes, report) = WriteMarkdown(section, options, cancellationToken);
        return OfficeConversionOutput.Save(report, path, bytes);
    }

    /// <summary>Asynchronously writes the section's Markdown projection as UTF-8.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsMarkdownAsync(this OneNoteSection section, string path, OneNoteMarkdownExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => WriteMarkdown(section, options, cancellationToken), cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, path, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Writes the section's Markdown projection as UTF-8 to a caller-owned stream and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsMarkdown(this OneNoteSection section, Stream stream, OneNoteMarkdownExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        var (bytes, report) = WriteMarkdown(section, options, cancellationToken);
        return OfficeConversionOutput.Save(report, stream, bytes);
    }

    /// <summary>Asynchronously writes the section's Markdown projection as UTF-8.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsMarkdownAsync(this OneNoteSection section, Stream stream, OneNoteMarkdownExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => WriteMarkdown(section, options, cancellationToken), cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, stream, bytes, cancellationToken).ConfigureAwait(false);
    }

    private static (byte[], OfficeConversionReport) WriteMarkdown(OneNoteSection section, OneNoteMarkdownExportOptions? options, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        OneNoteMarkdownConversionResult result = ConvertToMarkdownWithEvidence(section, options, out string text);
        if (options?.FailOnLoss == true) result.Report.ThrowIfLoss();
        return (OfficeConversionOutput.Text(text), result.Report);
    }

    /// <summary>Converts a notebook to a Markdown document.</summary>
    public static MarkdownDocument ToMarkdown(this OneNoteNotebook notebook, OneNoteMarkdownExportOptions? options = null, CancellationToken cancellationToken = default) =>
        notebook.ToMarkdown(options, out _, cancellationToken);

    /// <summary>Converts a notebook to a Markdown document and returns the completed conversion report.</summary>
    public static MarkdownDocument ToMarkdown(this OneNoteNotebook notebook, OneNoteMarkdownExportOptions? options, out OfficeConversionReport report,
        CancellationToken cancellationToken = default) {
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
        cancellationToken.ThrowIfCancellationRequested();
        OneNoteMarkdownConversionResult result = ConvertToMarkdownWithEvidence(notebook, options, out _);
        MarkdownDocument value = result.Value;
        OfficeConversionReport source = result.Report;
        return OfficeConversionOutput.Complete(value, source, options?.FailOnLoss == true, out report);
    }

    /// <summary>Converts a notebook to a Markdown document, returning the value and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static Task<OfficeConversionResult<MarkdownDocument>> ToMarkdownAsync(this OneNoteNotebook notebook, OneNoteMarkdownExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        return OfficeConversionOutput.RunAsync(() => (notebook.ToMarkdown(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);
    }

    /// <summary>Writes the notebook's Markdown projection as UTF-8 to a path and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsMarkdown(this OneNoteNotebook notebook, string path, OneNoteMarkdownExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        var (bytes, report) = WriteMarkdown(notebook, options, cancellationToken);
        return OfficeConversionOutput.Save(report, path, bytes);
    }

    /// <summary>Asynchronously writes the notebook's Markdown projection as UTF-8.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsMarkdownAsync(this OneNoteNotebook notebook, string path, OneNoteMarkdownExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => WriteMarkdown(notebook, options, cancellationToken), cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, path, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Writes the notebook's Markdown projection as UTF-8 to a caller-owned stream and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsMarkdown(this OneNoteNotebook notebook, Stream stream, OneNoteMarkdownExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        var (bytes, report) = WriteMarkdown(notebook, options, cancellationToken);
        return OfficeConversionOutput.Save(report, stream, bytes);
    }

    /// <summary>Asynchronously writes the notebook's Markdown projection as UTF-8.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsMarkdownAsync(this OneNoteNotebook notebook, Stream stream, OneNoteMarkdownExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => WriteMarkdown(notebook, options, cancellationToken), cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, stream, bytes, cancellationToken).ConfigureAwait(false);
    }

    private static (byte[], OfficeConversionReport) WriteMarkdown(OneNoteNotebook notebook, OneNoteMarkdownExportOptions? options, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        OneNoteMarkdownConversionResult result = ConvertToMarkdownWithEvidence(notebook, options, out string text);
        if (options?.FailOnLoss == true) result.Report.ThrowIfLoss();
        return (OfficeConversionOutput.Text(text), result.Report);
    }

}
