using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Markdown;

public partial class MarkdownDocument {
    /// <summary>Converts Markdown to HTML text using the requested fragment or document options.</summary>
    public string ToHtml(MarkdownHtmlExportOptions? options = null) => ToHtml(options, out _);

    /// <summary>Converts Markdown to HTML text and returns its fidelity report.</summary>
    public string ToHtml(MarkdownHtmlExportOptions? options, out OfficeConversionReport report) =>
        RenderHtmlConversion((options ?? new MarkdownHtmlExportOptions()).CloneForRender(), out report);

    /// <summary>Asynchronously converts Markdown to HTML text with its fidelity report.</summary>
    public async Task<OfficeConversionResult<string>> ToHtmlAsync(MarkdownHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = (options ?? new MarkdownHtmlExportOptions()).CloneForRender();
        return await Task.Run(() => {
            string html = ToHtml(options, out OfficeConversionReport report);
            cancellationToken.ThrowIfCancellationRequested();
            return new OfficeConversionResult<string>(html, report);
        }, cancellationToken).ConfigureAwait(false);
    }

    private string RenderHtmlConversion(MarkdownHtmlExportOptions operation, out OfficeConversionReport report) {
        var timer = Stopwatch.StartNew();
        var diagnostics = new List<OfficeDiagnostic>();
        Action<OfficeDiagnostic>? caller = operation.DiagnosticSink;
        operation.DiagnosticSink = diagnostic => { diagnostics.Add(diagnostic); caller?.Invoke(diagnostic); };
        string html = operation.Kind == MarkdownHtmlKind.Fragment && operation.IncludeCssInFragment
            ? RenderHtmlFragmentWithCss(operation) : HtmlRenderer.Render(this, operation);
        report = new OfficeConversionReport("Markdown", "Html", itemsWritten: Blocks.Count,
            elapsed: timer.Elapsed, diagnostics: diagnostics);
        if (operation.FailOnLoss) report.ThrowIfLoss();
        return html;
    }

    /// <summary>Writes HTML to a file, including an external CSS sidecar when requested.</summary>
    public OfficeConversionReport SaveAsHtml(string path, MarkdownHtmlExportOptions? options = null) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        var timer = Stopwatch.StartNew();
        MarkdownHtmlExportOptions operation = (options ?? new MarkdownHtmlExportOptions()).CloneForRender();
        if (operation.CssDelivery == MarkdownCssDelivery.ExternalFile) operation.ExternalCssOutputPath = Path.ChangeExtension(path, ".css");
        string html = RenderHtmlConversion(operation, out OfficeConversionReport report);
        global::OfficeIMO.Core.Internal.OfficeConversionFailures.Run(report, () => OfficeFileCommit.WriteAllBytes(path, Encoding.UTF8.GetBytes(html)));
        if (!string.IsNullOrEmpty(operation.ExternalCssOutputPath) && operation._externalCssContentToWrite != null)
            OfficeFileCommit.WriteAllBytes(operation.ExternalCssOutputPath!, Encoding.UTF8.GetBytes(operation._externalCssContentToWrite));
        return new OfficeConversionReport("Markdown", "Html", Path.GetFullPath(path), report.ItemsWritten, timer.Elapsed, report.Diagnostics);
    }

    /// <summary>Writes HTML to a caller-owned stream. External CSS requires a file destination.</summary>
    public OfficeConversionReport SaveAsHtml(Stream stream, MarkdownHtmlExportOptions? options = null) {
        if (options?.CssDelivery == MarkdownCssDelivery.ExternalFile) throw new ArgumentException("External CSS requires a file destination.", nameof(options));
        string html = ToHtml(options, out OfficeConversionReport report);
        global::OfficeIMO.Core.Internal.OfficeConversionFailures.Run(report, () => OfficeStreamWriter.WriteAllBytes(stream, Encoding.UTF8.GetBytes(html)));
        return report;
    }

    /// <summary>Asynchronously writes HTML to a file, including an external CSS sidecar when requested.</summary>
    public async Task<OfficeConversionReport> SaveAsHtmlAsync(string path, MarkdownHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = (options ?? new MarkdownHtmlExportOptions()).CloneForRender();
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        MarkdownHtmlExportOptions operation = (options ?? new MarkdownHtmlExportOptions()).CloneForRender();
        if (operation.CssDelivery == MarkdownCssDelivery.ExternalFile) operation.ExternalCssOutputPath = Path.ChangeExtension(path, ".css");
        string html = RenderHtmlConversion(operation, out OfficeConversionReport report);
        await global::OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(report, () => OfficeFileCommit.WriteAllBytesAsync(path, Encoding.UTF8.GetBytes(html), cancellationToken: cancellationToken)).ConfigureAwait(false);
        if (!string.IsNullOrEmpty(operation.ExternalCssOutputPath) && operation._externalCssContentToWrite != null)
            await OfficeFileCommit.WriteAllBytesAsync(operation.ExternalCssOutputPath!, Encoding.UTF8.GetBytes(operation._externalCssContentToWrite), cancellationToken: cancellationToken).ConfigureAwait(false);
        return new OfficeConversionReport("Markdown", "Html", Path.GetFullPath(path), report.ItemsWritten, timer.Elapsed, report.Diagnostics);
    }

    /// <summary>Asynchronously writes HTML to a caller-owned stream.</summary>
    public async Task<OfficeConversionReport> SaveAsHtmlAsync(Stream stream, MarkdownHtmlExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = (options ?? new MarkdownHtmlExportOptions()).CloneForRender();
        if (options?.CssDelivery == MarkdownCssDelivery.ExternalFile) throw new ArgumentException("External CSS requires a file destination.", nameof(options));
        cancellationToken.ThrowIfCancellationRequested();
        string html = ToHtml(options, out OfficeConversionReport report);
        await global::OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(report, () => OfficeStreamWriter.WriteAllBytesAsync(stream, Encoding.UTF8.GetBytes(html), cancellationToken)).ConfigureAwait(false);
        return report;
    }
}
