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
        return new AdfConversionResult<MarkdownDoc>(MarkdownReader.Parse(result.Value), result.Report.Diagnostics);
    }

    /// <summary>Converts an ADF document to Markdown text.</summary>
    public static AdfConversionResult<string> ToMarkdown(
        AdfDocument document,
        AdfConversionOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        var diagnostics = new List<AdfConversionDiagnostic>();
        string markdown = AdfToMarkdownConverter.Convert(document, options ?? new AdfConversionOptions(), diagnostics);
        return new AdfConversionResult<string>(markdown, diagnostics);
    }

    /// <summary>Converts an ADF document to an HTML fragment through OfficeIMO.Markdown.</summary>
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
        return new AdfConversionResult<string>(result.Value.ToHtmlFragment(htmlOptions), diagnostics);
    }

    /// <summary>Converts Markdown text to ADF.</summary>
    public static AdfConversionResult<AdfDocument> FromMarkdown(string markdown) {
        if (markdown == null) throw new ArgumentNullException(nameof(markdown));
        return FromMarkdown(MarkdownReader.Parse(markdown));
    }

    /// <summary>Converts an OfficeIMO Markdown document to ADF.</summary>
    public static AdfConversionResult<AdfDocument> FromMarkdown(MarkdownDoc markdown) {
        if (markdown == null) throw new ArgumentNullException(nameof(markdown));
        var diagnostics = new List<AdfConversionDiagnostic>();
        AdfDocument value = MarkdownToAdfConverter.Convert(markdown, diagnostics);
        return new AdfConversionResult<AdfDocument>(value, diagnostics);
    }

    /// <summary>Converts HTML to ADF through OfficeIMO.Html and OfficeIMO.Markdown.Html.</summary>
    public static AdfConversionResult<AdfDocument> FromHtml(string html, HtmlToMarkdownOptions? options = null) {
        if (html == null) throw new ArgumentNullException(nameof(html));
        MarkdownDoc markdown = HtmlConversionDocument.Parse(html).ToMarkdownDocument(options);
        AdfConversionResult<AdfDocument> result = FromMarkdown(markdown);
        var diagnostics = result.Report.Diagnostics.ToList();
        diagnostics.Insert(0, new AdfConversionDiagnostic(
            "ADF_HTML_VIA_MARKDOWN",
            "$",
            "HTML is projected through the OfficeIMO Markdown model before ADF generation.",
            AdfConversionSeverity.Warning));
        return new AdfConversionResult<AdfDocument>(result.Value, diagnostics);
    }
}
