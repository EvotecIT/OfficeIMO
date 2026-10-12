using OfficeIMO.Markdown.Html;
using OfficeIMO.Markdown;

namespace OfficeIMO.Chm;

/// <summary>Converts compiled help through the canonical HTML-to-Markdown engine.</summary>
public static class ChmMarkdownConverterExtensions {
    /// <summary>Exports selected topics in book order, retaining linked topic anchors and reporting projection loss.</summary>
    public static ChmConversionResult<string> ToMarkdownResult(this ChmDocument document, ChmConversionOptions? options = null,
        HtmlToMarkdownOptions? markdownOptions = null, CancellationToken cancellationToken = default) =>
        Convert(document, options, markdownOptions, cancellationToken, (_, text) => text);

    /// <summary>Exports an editable canonical Markdown document with the same selection, output budget and fidelity evidence.</summary>
    public static ChmConversionResult<MarkdownDoc> ToMarkdownDocumentResult(this ChmDocument document, ChmConversionOptions? options = null,
        HtmlToMarkdownOptions? markdownOptions = null, CancellationToken cancellationToken = default) =>
        Convert(document, options, markdownOptions, cancellationToken, (value, _) => value);

    private static ChmConversionResult<T> Convert<T>(ChmDocument document, ChmConversionOptions? options,
        HtmlToMarkdownOptions? markdownOptions, CancellationToken cancellationToken, Func<MarkdownDoc, string, T> select) where T : class {
        if (document == null) throw new ArgumentNullException(nameof(document));
        ChmConversionOptions configured = options?.Clone() ?? new ChmConversionOptions();
        ChmConversionResult<HtmlConversionDocument> projected = document.ToHtmlDocumentResult(configured, cancellationToken);
        HtmlToMarkdownOptions conversion = markdownOptions?.Clone() ?? HtmlToMarkdownOptions.CreatePortableProfile();
        conversion.MaxInputCharacters = Math.Min(conversion.MaxInputCharacters ?? int.MaxValue, (int)configured.MaxOutputBytes);
        var converted = projected.Value.ToMarkdownDocumentResult(conversion, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        string markdown = converted.Value.ToMarkdown(conversion.MarkdownWriteOptions, cancellationToken);
        ChmDocument.EnforceOutput(Encoding.UTF8.GetByteCount(markdown), configured);
        cancellationToken.ThrowIfCancellationRequested();
        var report = new ChmConversionReport(projected.Report.TopicPaths,
            projected.Report.FidelityDiagnostics.Concat(converted.Report.FidelityDiagnostics));
        return new ChmConversionResult<T>(select(converted.Value, markdown), report);
    }

    /// <summary>Exports a linked Markdown book. Use <see cref="ToMarkdownResult"/> to inspect fidelity evidence.</summary>
    public static string ToMarkdown(this ChmDocument document, ChmConversionOptions? options = null,
        HtmlToMarkdownOptions? markdownOptions = null, CancellationToken cancellationToken = default) =>
        document.ToMarkdownResult(options, markdownOptions, cancellationToken).RequireValue();
}
