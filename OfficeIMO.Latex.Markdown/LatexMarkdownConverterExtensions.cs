using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Latex.Markdown;

/// <summary>Loss-aware conversions between Latex and Markdown models.</summary>
public static class LatexMarkdownConverterExtensions {
    /// <summary>Converts into a detached Markdown model.</summary>
    public static MarkdownDocument ToMarkdown(this LatexDocument document, LatexMarkdownExportOptions? options = null,
        CancellationToken cancellationToken = default) => document.ToMarkdown(options, out _, cancellationToken);

    /// <summary>Converts and retains native diagnostic locations and canonical source evidence.</summary>
    public static MarkdownDocument ToMarkdown(this LatexDocument document, LatexMarkdownExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var result = LatexToMarkdownConverter.Project(document, options, cancellationToken).Result;
        cancellationToken.ThrowIfCancellationRequested();
        return OfficeConversionOutput.Complete(result.Value, result.Report, options?.FailOnLoss == true, out report);
    }

    /// <summary>Converts asynchronously and returns the model and completed report.</summary>
    public static Task<OfficeConversionResult<MarkdownDocument>> ToMarkdownAsync(this LatexDocument document,
        LatexMarkdownExportOptions? options = null, CancellationToken cancellationToken = default) =>
        OfficeConversionOutput.RunAsync(() => (document.ToMarkdown(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);

    /// <summary>Writes converted Markdown and returns the report, including native evidence.</summary>
    public static OfficeConversionReport SaveAsMarkdown(this LatexDocument document,
        string path, LatexMarkdownExportOptions? options = null, CancellationToken cancellationToken = default) {
        MarkdownDocument value = document.ToMarkdown(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return OfficeConversionOutput.Save(report, path, OfficeConversionOutput.Text(value.ToMarkdown()));
    }

    /// <summary>Writes converted Markdown and returns the report, including native evidence.</summary>
    public static OfficeConversionReport SaveAsMarkdown(this LatexDocument document,
        Stream stream, LatexMarkdownExportOptions? options = null, CancellationToken cancellationToken = default) {
        MarkdownDocument value = document.ToMarkdown(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        OfficeConversionOutput.Save(report, stream, OfficeConversionOutput.Text(value.ToMarkdown()));
        return report;
    }

    /// <summary>Asynchronously writes converted Markdown and returns the report, including native evidence.</summary>
    public static async Task<OfficeConversionReport> SaveAsMarkdownAsync(this LatexDocument document,
        string path, LatexMarkdownExportOptions? options = null, CancellationToken cancellationToken = default) {
        MarkdownDocument value = document.ToMarkdown(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return await OfficeConversionOutput.SaveAsync(report, path, OfficeConversionOutput.Text(value.ToMarkdown()), cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes converted Markdown and returns the report, including native evidence.</summary>
    public static async Task<OfficeConversionReport> SaveAsMarkdownAsync(this LatexDocument document,
        Stream stream, LatexMarkdownExportOptions? options = null, CancellationToken cancellationToken = default) {
        MarkdownDocument value = document.ToMarkdown(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        await OfficeConversionOutput.SaveAsync(report, stream, OfficeConversionOutput.Text(value.ToMarkdown()), cancellationToken).ConfigureAwait(false);
        return report;
    }

    /// <summary>Converts into a detached Latex model.</summary>
    public static LatexDocument ToLatex(this MarkdownDocument document, MarkdownLatexExportOptions? options = null,
        CancellationToken cancellationToken = default) => document.ToLatex(options, out _, cancellationToken);

    /// <summary>Converts and retains native diagnostic locations and canonical source evidence.</summary>
    public static LatexDocument ToLatex(this MarkdownDocument document, MarkdownLatexExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var result = MarkdownToLatexConverter.Convert(document, options, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return OfficeConversionOutput.Complete(result.Value, result.Report, options?.FailOnLoss == true, out report);
    }

    /// <summary>Converts asynchronously and returns the model and completed report.</summary>
    public static Task<OfficeConversionResult<LatexDocument>> ToLatexAsync(this MarkdownDocument document,
        MarkdownLatexExportOptions? options = null, CancellationToken cancellationToken = default) =>
        OfficeConversionOutput.RunAsync(() => (document.ToLatex(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);

    /// <summary>Writes converted Latex and returns the report, including native evidence.</summary>
    public static OfficeConversionReport SaveAsLatex(this MarkdownDocument document,
        string path, MarkdownLatexExportOptions? options = null, CancellationToken cancellationToken = default) {
        LatexDocument value = document.ToLatex(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return OfficeConversionOutput.Save(report, path, OfficeConversionOutput.Text(((MarkdownToLatexReport)report).GeneratedSource));
    }

    /// <summary>Writes converted Latex and returns the report, including native evidence.</summary>
    public static OfficeConversionReport SaveAsLatex(this MarkdownDocument document,
        Stream stream, MarkdownLatexExportOptions? options = null, CancellationToken cancellationToken = default) {
        LatexDocument value = document.ToLatex(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        OfficeConversionOutput.Save(report, stream, OfficeConversionOutput.Text(((MarkdownToLatexReport)report).GeneratedSource));
        return report;
    }

    /// <summary>Asynchronously writes converted Latex and returns the report, including native evidence.</summary>
    public static async Task<OfficeConversionReport> SaveAsLatexAsync(this MarkdownDocument document,
        string path, MarkdownLatexExportOptions? options = null, CancellationToken cancellationToken = default) {
        LatexDocument value = document.ToLatex(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return await OfficeConversionOutput.SaveAsync(report, path, OfficeConversionOutput.Text(((MarkdownToLatexReport)report).GeneratedSource), cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes converted Latex and returns the report, including native evidence.</summary>
    public static async Task<OfficeConversionReport> SaveAsLatexAsync(this MarkdownDocument document,
        Stream stream, MarkdownLatexExportOptions? options = null, CancellationToken cancellationToken = default) {
        LatexDocument value = document.ToLatex(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        await OfficeConversionOutput.SaveAsync(report, stream, OfficeConversionOutput.Text(((MarkdownToLatexReport)report).GeneratedSource), cancellationToken).ConfigureAwait(false);
        return report;
    }

}
