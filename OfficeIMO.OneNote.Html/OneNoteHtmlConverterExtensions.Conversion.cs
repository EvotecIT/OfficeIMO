using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;
using OfficeIMO.Html;

namespace OfficeIMO.OneNote.Html;

public static partial class OneNoteHtmlConverterExtensions {
    /// <summary>Converts a section to semantic HTML through the shared Markdown model.</summary>
    public static string ToHtml(this OneNoteSection section, OneNoteHtmlExportOptions? options = null, CancellationToken cancellationToken = default) =>
        section.ToHtml(options, out _, cancellationToken);

    /// <summary>Converts a section to semantic HTML through the shared Markdown model and returns the completed conversion report.</summary>
    public static string ToHtml(this OneNoteSection section, OneNoteHtmlExportOptions? options, out OfficeConversionReport report,
        CancellationToken cancellationToken = default) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        cancellationToken.ThrowIfCancellationRequested();
        HtmlTextConversionResult result = section.ConvertToHtmlWithEvidence(options);
        string value = result.RequireValue();
        OfficeConversionReport source = HtmlConversionEvidence.CreateReport("OneNote", "Html", result.Report.NativeDiagnostics);
        return OfficeConversionOutput.Complete(value, source, options?.FailOnLoss == true, out report);
    }

    /// <summary>Converts a section to semantic HTML through the shared Markdown model, returning the value and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static Task<OfficeConversionResult<string>> ToHtmlAsync(this OneNoteSection section, OneNoteHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        return OfficeConversionOutput.RunAsync(() => (section.ToHtml(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);
    }

    /// <summary>Writes the converted output to a path and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsHtml(this OneNoteSection section, string path, OneNoteHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        string value = section.ToHtml(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, path, OfficeConversionOutput.Text(value));
    }

    /// <summary>Writes the converted output to a caller-owned stream, which stays open and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsHtml(this OneNoteSection section, Stream stream, OneNoteHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        string value = section.ToHtml(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, stream, OfficeConversionOutput.Text(value));
    }

    /// <summary>Asynchronously writes the converted output to a path and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsHtmlAsync(this OneNoteSection section, string path, OneNoteHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            string value = section.ToHtml(options, out OfficeConversionReport converted, cancellationToken);
            return (OfficeConversionOutput.Text(value), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, path, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes the converted output to a caller-owned stream and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsHtmlAsync(this OneNoteSection section, Stream stream, OneNoteHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            string value = section.ToHtml(options, out OfficeConversionReport converted, cancellationToken);
            return (OfficeConversionOutput.Text(value), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, stream, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Converts a notebook to semantic HTML through the shared Markdown model.</summary>
    public static string ToHtml(this OneNoteNotebook notebook, OneNoteHtmlExportOptions? options = null, CancellationToken cancellationToken = default) =>
        notebook.ToHtml(options, out _, cancellationToken);

    /// <summary>Converts a notebook to semantic HTML through the shared Markdown model and returns the completed conversion report.</summary>
    public static string ToHtml(this OneNoteNotebook notebook, OneNoteHtmlExportOptions? options, out OfficeConversionReport report,
        CancellationToken cancellationToken = default) {
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
        cancellationToken.ThrowIfCancellationRequested();
        HtmlTextConversionResult result = notebook.ConvertToHtmlWithEvidence(options);
        string value = result.RequireValue();
        OfficeConversionReport source = HtmlConversionEvidence.CreateReport("OneNote", "Html", result.Report.NativeDiagnostics);
        return OfficeConversionOutput.Complete(value, source, options?.FailOnLoss == true, out report);
    }

    /// <summary>Converts a notebook to semantic HTML through the shared Markdown model, returning the value and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static Task<OfficeConversionResult<string>> ToHtmlAsync(this OneNoteNotebook notebook, OneNoteHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        return OfficeConversionOutput.RunAsync(() => (notebook.ToHtml(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);
    }

    /// <summary>Writes the converted output to a path and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsHtml(this OneNoteNotebook notebook, string path, OneNoteHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        string value = notebook.ToHtml(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, path, OfficeConversionOutput.Text(value));
    }

    /// <summary>Writes the converted output to a caller-owned stream, which stays open and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsHtml(this OneNoteNotebook notebook, Stream stream, OneNoteHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        string value = notebook.ToHtml(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, stream, OfficeConversionOutput.Text(value));
    }

    /// <summary>Asynchronously writes the converted output to a path and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsHtmlAsync(this OneNoteNotebook notebook, string path, OneNoteHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            string value = notebook.ToHtml(options, out OfficeConversionReport converted, cancellationToken);
            return (OfficeConversionOutput.Text(value), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, path, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes the converted output to a caller-owned stream and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsHtmlAsync(this OneNoteNotebook notebook, Stream stream, OneNoteHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            string value = notebook.ToHtml(options, out OfficeConversionReport converted, cancellationToken);
            return (OfficeConversionOutput.Text(value), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, stream, bytes, cancellationToken).ConfigureAwait(false);
    }

}
