using System.Diagnostics;
using OfficeIMO.Markdown;

namespace OfficeIMO.Notebook.Markdown;

/// <summary>Projects static content from a Jupyter notebook to the typed Markdown model.</summary>
public static class NotebookMarkdownConverterExtensions {
    /// <summary>Projects prose, code and stored output according to the selected presentation options.</summary>
    public static MarkdownDocument ToMarkdown(this NotebookDocument notebook, NotebookMarkdownExportOptions? options = null,
        CancellationToken cancellationToken = default) => notebook.ToMarkdown(options, out _, cancellationToken);

    /// <summary>Projects the notebook and reports fidelity findings at the source cell or output.</summary>
    public static MarkdownDocument ToMarkdown(this NotebookDocument notebook, NotebookMarkdownExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(notebook);
        cancellationToken.ThrowIfCancellationRequested();
        NotebookMarkdownExportOptions operation = (options ?? new NotebookMarkdownExportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        var projection = new NotebookMarkdownProjection(notebook, operation, cancellationToken);
        MarkdownDocument result = projection.Project();
        report = new OfficeConversionReport("Jupyter", "Markdown", itemsWritten: result.Blocks.Count,
            elapsed: timer.Elapsed, diagnostics: projection.Diagnostics);
        if (operation.FailOnLoss && report.HasLoss) {
            result.Dispose();
            report.ThrowIfLoss();
        }
        return result;
    }
}
