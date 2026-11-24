using System.Diagnostics;
using OfficeIMO.Core.Internal;
using OfficeIMO.Markdown;

namespace OfficeIMO.Notebook.Markdown;

/// <summary>Projects static content from a Jupyter notebook to the typed Markdown model.</summary>
public static class NotebookMarkdownConverterExtensions {
    /// <summary>Projects prose, code and stored output according to the selected presentation options.</summary>
    public static MarkdownDocument ToMarkdown(this NotebookDocument notebook, NotebookMarkdownExportOptions? options = null,
        CancellationToken cancellationToken = default) => notebook.ToMarkdown(options, out _, cancellationToken);

    /// <summary>Projects the notebook and reports fidelity findings at the source cell or output.</summary>
    public static MarkdownDocument ToMarkdown(this NotebookDocument notebook, NotebookMarkdownExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) =>
        Project(notebook, options, preserveSource: false, out report, cancellationToken);

    /// <summary>Returns Markdown text, optionally preserving the exact source within Markdown cells.</summary>
    public static string ToMarkdownText(this NotebookDocument notebook, NotebookMarkdownExportOptions? options = null,
        CancellationToken cancellationToken = default) => notebook.ToMarkdownText(options, out _, cancellationToken);

    /// <summary>Writes Markdown text, optionally preserving the exact source within Markdown cells.</summary>
    public static string ToMarkdownText(this NotebookDocument notebook, NotebookMarkdownExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        using MarkdownDocument projected = Project(notebook, options, preserveSource: options?.PreserveMarkdownSource == true,
            out report, cancellationToken);
        return projected.ToMarkdown(options: null, cancellationToken);
    }

    /// <summary>Writes UTF-8 Markdown to a caller-owned stream and reports conversion findings.</summary>
    public static async Task<OfficeConversionReport> SaveAsMarkdownAsync(this NotebookDocument notebook, Stream destination,
        NotebookMarkdownExportOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(destination);
        string content = notebook.ToMarkdownText(options, out OfficeConversionReport report, cancellationToken);
        return await OfficeConversionOutput.SaveAsync(report, destination, OfficeConversionOutput.Text(content), cancellationToken)
            .ConfigureAwait(false);
    }

    /// <summary>Writes UTF-8 Markdown to a path and reports conversion findings.</summary>
    public static async Task<OfficeConversionReport> SaveAsMarkdownAsync(this NotebookDocument notebook, string path,
        NotebookMarkdownExportOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        string content = notebook.ToMarkdownText(options, out OfficeConversionReport report, cancellationToken);
        return await OfficeConversionOutput.SaveAsync(report, path, OfficeConversionOutput.Text(content), cancellationToken)
            .ConfigureAwait(false);
    }

    private static MarkdownDocument Project(NotebookDocument notebook, NotebookMarkdownExportOptions? options, bool preserveSource,
        out OfficeConversionReport report, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(notebook);
        cancellationToken.ThrowIfCancellationRequested();
        NotebookMarkdownExportOptions operation = options?.Clone() ?? new NotebookMarkdownExportOptions();
        var timer = Stopwatch.StartNew();
        var projection = new NotebookMarkdownProjection(notebook, operation, cancellationToken);
        MarkdownDocument result = projection.Project(preserveSource);
        report = new OfficeConversionReport("Jupyter", "Markdown", itemsWritten: result.Blocks.Count,
            elapsed: timer.Elapsed, diagnostics: projection.Diagnostics);
        if (operation.FailOnLoss && report.HasLoss) {
            result.Dispose();
            report.ThrowIfLoss();
        }
        return result;
    }
}
