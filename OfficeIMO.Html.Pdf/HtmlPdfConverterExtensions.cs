using OfficeIMO.Markdown.Html;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Pdf;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

/// <summary>
/// First-party HTML to PDF conversion helpers.
/// </summary>
public static class HtmlPdfConverterExtensions {
    /// <summary>
    /// Converts HTML text to a first-party OfficeIMO PDF document model.
    /// </summary>
    public static PdfCore.PdfDocument ToPdfDocument(this string html, HtmlPdfSaveOptions? options = null) {
        if (html == null) {
            throw new ArgumentNullException(nameof(html));
        }

        options ??= new HtmlPdfSaveOptions();
        options.ResetExportState();

        return options.Profile switch {
            HtmlPdfProfile.Semantic => ConvertSemantic(html, options),
            HtmlPdfProfile.Document => ConvertDocument(html, options),
            _ => throw new ArgumentOutOfRangeException(nameof(options.Profile), options.Profile, "Unsupported HTML PDF profile.")
        };
    }

    /// <summary>
    /// Converts HTML stream content to a first-party OfficeIMO PDF document model using UTF-8.
    /// </summary>
    public static PdfCore.PdfDocument ToPdfDocument(this Stream htmlStream, HtmlPdfSaveOptions? options = null) {
        if (htmlStream == null) {
            throw new ArgumentNullException(nameof(htmlStream));
        }

        using var reader = new StreamReader(htmlStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 4096, leaveOpen: true);
        return reader.ReadToEnd().ToPdfDocument(options);
    }

    /// <summary>
    /// Converts HTML text to PDF bytes.
    /// </summary>
    public static byte[] SaveAsPdf(this string html, HtmlPdfSaveOptions? options = null) {
        return html.ToPdfDocument(options).ToBytes();
    }

    /// <summary>
    /// Saves HTML text as a PDF file.
    /// </summary>
    public static void SaveAsPdf(this string html, string path, HtmlPdfSaveOptions? options = null) {
        html.ToPdfDocument(options).Save(path);
    }

    /// <summary>
    /// Attempts to save HTML text as a PDF file and returns output diagnostics instead of throwing.
    /// </summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this string html, string path, HtmlPdfSaveOptions? options = null) {
        try {
            return html.ToPdfDocument(options).TrySave(path);
        } catch (Exception ex) {
            return PdfCore.PdfSaveResult.FromFailure(path, ex);
        }
    }

    /// <summary>
    /// Writes HTML text as PDF to a stream.
    /// </summary>
    public static void SaveAsPdf(this string html, Stream stream, HtmlPdfSaveOptions? options = null) {
        html.ToPdfDocument(options).Save(stream);
    }

    /// <summary>
    /// Attempts to write HTML text as PDF to a stream and returns output diagnostics instead of throwing.
    /// </summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this string html, Stream stream, HtmlPdfSaveOptions? options = null) {
        try {
            return html.ToPdfDocument(options).TrySave(stream);
        } catch (Exception ex) {
            return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex);
        }
    }

    private static PdfCore.PdfDocument ConvertSemantic(string html, HtmlPdfSaveOptions options) {
        MarkdownPdfSaveOptions markdownPdfOptions = options.MarkdownPdfOptions ?? new MarkdownPdfSaveOptions();
        PdfCore.PdfDocument pdf = html
            .LoadFromHtml(options.MarkdownHtmlOptions)
            .ToPdfDocument(markdownPdfOptions);
        options.ConversionReport.AddRange(markdownPdfOptions.ConversionReport.Warnings);
        return pdf;
    }

    private static PdfCore.PdfDocument ConvertDocument(string html, HtmlPdfSaveOptions options) {
        PdfSaveOptions wordPdfOptions = options.WordPdfOptions ?? new PdfSaveOptions();
        using WordDocument document = html.LoadFromHtml(options.WordHtmlOptions);
        PdfCore.PdfDocument pdf = document.ToPdfDocument(wordPdfOptions);
        options.ConversionReport.AddRange(wordPdfOptions.ConversionReport.Warnings);
        return pdf;
    }
}
