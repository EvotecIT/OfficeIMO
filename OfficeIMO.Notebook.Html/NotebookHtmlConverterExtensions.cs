using System.Diagnostics;
using OfficeIMO.Core.Internal;
using OfficeIMO.Markdown;
using OfficeIMO.Notebook.Markdown;

namespace OfficeIMO.Notebook.Html;

/// <summary>Converts saved notebook cells to offline HTML documents or fragments.</summary>
public static class NotebookHtmlConverterExtensions {
    /// <summary>Converts the notebook to offline HTML.</summary>
    public static string ToHtml(this NotebookDocument notebook, NotebookHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) => notebook.ToHtml(options, out _, cancellationToken);

    /// <summary>Converts the notebook to offline HTML and returns its fidelity report.</summary>
    public static string ToHtml(this NotebookDocument notebook, NotebookHtmlExportOptions? options, out OfficeConversionReport report,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(notebook);
        cancellationToken.ThrowIfCancellationRequested();
        NotebookHtmlExportOptions operation = (options ?? new NotebookHtmlExportOptions()).Clone();
        bool failOnLoss = operation.FailOnLoss || operation.ProjectionOptions.FailOnLoss;
        operation.ProjectionOptions.FailOnLoss = false;
        var timer = Stopwatch.StartNew();
        NotebookHtmlWidgets? widgets = operation.IncludeInteractiveWidgets && operation.ProjectionOptions.IncludeOutputs
            ? new NotebookHtmlWidgets(notebook, cancellationToken)
            : null;
        if (widgets != null) operation.ProjectionOptions.OutputRenderer = widgets.RenderOutput;
        using MarkdownDocument projection = notebook.ToMarkdown(operation.ProjectionOptions, out OfficeConversionReport projectionReport, cancellationToken);
        var htmlOptions = new MarkdownHtmlExportOptions {
            Kind = operation.Fragment ? MarkdownHtmlKind.Fragment : MarkdownHtmlKind.Document,
            Title = operation.ProjectionOptions.Title ?? "Notebook",
            Style = operation.Style,
            AssetMode = MarkdownAssetMode.Offline,
            RawHtmlHandling = MarkdownRawHtmlHandling.Escape,
            BlockExternalHttpImages = true,
            AllowTrustedSvgDataImages = true
        };
        var mathDiagnostics = new List<OfficeDiagnostic>();
        NotebookHtmlMath.Register(htmlOptions, mathDiagnostics, cancellationToken);
        string html = projection.ToHtml(htmlOptions, out OfficeConversionReport htmlReport);
        if (widgets != null) html = widgets.AddRuntime(html, operation.Fragment);
        cancellationToken.ThrowIfCancellationRequested();
        var findings = new OfficeConversionReport("Jupyter", "Html", itemsWritten: projection.Blocks.Count,
            elapsed: timer.Elapsed, diagnostics: projectionReport.Diagnostics.Concat(htmlReport.Diagnostics).Concat(mathDiagnostics)
                .Concat(widgets?.Diagnostics ?? Array.Empty<OfficeDiagnostic>()));
        return OfficeConversionOutput.Complete(html, findings, failOnLoss, out report);
    }

    /// <summary>Converts the notebook to offline HTML and returns its content and report.</summary>
    public static Task<OfficeConversionResult<string>> ToHtmlAsync(this NotebookDocument notebook, NotebookHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        return OfficeConversionOutput.RunAsync(() => (notebook.ToHtml(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);
    }

    /// <summary>Writes UTF-8 HTML to a path and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsHtml(this NotebookDocument notebook, string path, NotebookHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        string value = notebook.ToHtml(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, path, OfficeConversionOutput.Text(value));
    }

    /// <summary>Writes UTF-8 HTML to a caller-owned stream and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsHtml(this NotebookDocument notebook, Stream stream, NotebookHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(stream);
        string value = notebook.ToHtml(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, stream, OfficeConversionOutput.Text(value));
    }

    /// <summary>Asynchronously writes UTF-8 HTML to a path and returns the conversion report.</summary>
    public static async Task<OfficeConversionReport> SaveAsHtmlAsync(this NotebookDocument notebook, string path, NotebookHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        OfficeConversionResult<string> converted = await notebook.ToHtmlAsync(options, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(converted.Report, path, OfficeConversionOutput.Text(converted.Value), cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes UTF-8 HTML to a caller-owned stream and returns the conversion report.</summary>
    public static async Task<OfficeConversionReport> SaveAsHtmlAsync(this NotebookDocument notebook, Stream stream, NotebookHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(stream);
        OfficeConversionResult<string> converted = await notebook.ToHtmlAsync(options, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(converted.Report, stream, OfficeConversionOutput.Text(converted.Value), cancellationToken).ConfigureAwait(false);
    }
}
