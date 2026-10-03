using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;

namespace OfficeIMO.Adf;

/// <summary>Converts ADF through OfficeIMO's canonical Markdown and HTML models.</summary>
public static class AdfConverter {
    /// <summary>Converts an ADF document to the OfficeIMO Markdown object model.</summary>
    public static AdfConversionResult<MarkdownDoc> ToMarkdownDocument(
        AdfDocument document,
        AdfConversionOptions? options = null) {
        AdfConversionResult<string> result = ToMarkdown(document, options);
        options ??= new AdfConversionOptions();
        MarkdownDoc value = MarkdownReader.Parse(result.Value, new MarkdownReaderOptions { MaxInputCharacters = options.MaxOutputCharacters, MaxNestingDepth = options.MaxDepth });
        options.CancellationToken.ThrowIfCancellationRequested();
        return new AdfConversionResult<MarkdownDoc>(value, result.Report.Diagnostics);
    }

    /// <summary>Converts an ADF document to Markdown text.</summary>
    public static AdfConversionResult<string> ToMarkdown(
        AdfDocument document,
        AdfConversionOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        options ??= new AdfConversionOptions();
        AdfGraphGuard.Check(document, options);
        var diagnostics = new List<AdfConversionDiagnostic>();
        string markdown = AdfToMarkdownConverter.Convert(document, options, diagnostics);
        AdfGraphGuard.Output(markdown, options);
        return new AdfConversionResult<string>(markdown, diagnostics);
    }

    /// <summary>Converts an ADF document to an HTML fragment through OfficeIMO.Markdown.</summary>
    /// <remarks>The report always includes a warning for the intermediate Markdown projection, even if no other fidelity issue is found.</remarks>
    public static AdfConversionResult<string> ToHtml(
        AdfDocument document,
        HtmlOptions? htmlOptions = null,
        AdfConversionOptions? options = null) {
        AdfConversionResult<MarkdownDoc> result = ToMarkdownDocument(document, options);
        var diagnostics = result.Report.Diagnostics.ToList();
        diagnostics.Add(new AdfConversionDiagnostic(
            "ADF_TO_HTML_VIA_MARKDOWN",
            "$",
            "ADF is projected through the OfficeIMO Markdown model before HTML rendering.",
            AdfConversionSeverity.Warning));
        string html = result.Value.ToHtmlFragment(htmlOptions);
        AdfGraphGuard.Output(html, options ?? new AdfConversionOptions());
        return new AdfConversionResult<string>(html, diagnostics);
    }

    /// <summary>Converts Markdown text to ADF.</summary>
    public static AdfConversionResult<AdfDocument> FromMarkdown(string markdown) {
        return FromMarkdown(markdown, new AdfConversionOptions());
    }

    /// <summary>Converts Markdown text to ADF with resource limits and cancellation.</summary>
    public static AdfConversionResult<AdfDocument> FromMarkdown(string markdown, AdfConversionOptions options) {
        if (markdown == null) throw new ArgumentNullException(nameof(markdown));
        if (options == null) throw new ArgumentNullException(nameof(options));
        AdfGraphGuard.Input(markdown, options);
        return FromMarkdown(MarkdownReader.Parse(markdown, new MarkdownReaderOptions { MaxInputCharacters = markdown.Length, MaxNestingDepth = options.MaxDepth }), options);
    }

    /// <summary>Converts an OfficeIMO Markdown document to ADF.</summary>
    public static AdfConversionResult<AdfDocument> FromMarkdown(MarkdownDoc markdown) {
        return FromMarkdown(markdown, new AdfConversionOptions());
    }

    /// <summary>Converts an OfficeIMO Markdown document to ADF with resource limits and cancellation.</summary>
    public static AdfConversionResult<AdfDocument> FromMarkdown(MarkdownDoc markdown, AdfConversionOptions options) {
        if (markdown == null) throw new ArgumentNullException(nameof(markdown));
        if (options == null) throw new ArgumentNullException(nameof(options));
        AdfGraphGuard.CheckMarkdown(markdown, options);
        var diagnostics = new List<AdfConversionDiagnostic>();
        AdfDocument value = MarkdownToAdfConverter.Convert(markdown, diagnostics);
        AdfGraphGuard.Check(value, options);
        return new AdfConversionResult<AdfDocument>(value, diagnostics);
    }

    /// <summary>Converts HTML to ADF through OfficeIMO.Html and OfficeIMO.Markdown.Html.</summary>
    /// <remarks>The report always includes a warning for the intermediate Markdown projection.</remarks>
    public static AdfConversionResult<AdfDocument> FromHtml(string html, HtmlToMarkdownOptions? options = null) {
        return FromHtml(html, options, new AdfConversionOptions());
    }

    /// <summary>Converts HTML to ADF with resource limits and cancellation around the canonical HTML/Markdown bridge.</summary>
    public static AdfConversionResult<AdfDocument> FromHtml(string html, HtmlToMarkdownOptions? options, AdfConversionOptions processingOptions) {
        if (html == null) throw new ArgumentNullException(nameof(html));
        if (processingOptions == null) throw new ArgumentNullException(nameof(processingOptions));
        AdfGraphGuard.Input(html, processingOptions);
        MarkdownDoc markdown = HtmlConversionDocument.Parse(html).ToMarkdownDocument(options);
        AdfConversionResult<AdfDocument> result = FromMarkdown(markdown, processingOptions);
        var diagnostics = result.Report.Diagnostics.ToList();
        diagnostics.Insert(0, new AdfConversionDiagnostic(
            "ADF_HTML_VIA_MARKDOWN",
            "$",
            "HTML is projected through the OfficeIMO Markdown model before ADF generation.",
            AdfConversionSeverity.Warning));
        return new AdfConversionResult<AdfDocument>(result.Value, diagnostics);
    }
}
