using OfficeIMO.Web.Converter.Models;

namespace OfficeIMO.Web.Converter.Services;

public static class ConversionRouteCatalog {
    public static IReadOnlyList<ConversionRoute> All { get; } = [
        new("docx-pdf", "DOCX", "PDF", "Word to PDF", "Convert a Word document into a downloadable PDF.", ConversionInputKind.File, ".docx", "WordDocument.Load(stream).ToPdfDocumentResult(options)", "ocx-route-card--word"),
        new("xlsx-pdf", "XLSX", "PDF", "Excel to PDF", "Render workbook sheets with layout and conversion diagnostics.", ConversionInputKind.File, ".xlsx", "ExcelDocument.Load(stream).ToPdfDocumentResult(options)", "ocx-route-card--excel"),
        new("pptx-pdf", "PPTX", "PDF", "PowerPoint to PDF", "Render presentation slides into a portable PDF.", ConversionInputKind.File, ".pptx", "PowerPointPresentation.Load(stream).ToPdfDocumentResult(options)", "ocx-route-card--powerpoint"),
        new("markdown-html", "MD", "HTML", "Markdown to HTML", "Render Markdown into an immediate browser preview and HTML download.", ConversionInputKind.Text, ".md,.markdown,.txt", "MarkdownRenderer.RenderBodyHtml(markdown, options)", "ocx-route-card--markdown"),
        new("html-markdown", "HTML", "MD", "HTML to Markdown", "Turn HTML into portable Markdown with a shared resource policy.", ConversionInputKind.Text, ".html,.htm,.txt", "HtmlConversionDocument.Parse(html).ToMarkdown(options)", "ocx-route-card--html"),
        new("markdown-docx", "MD", "DOCX", "Markdown to Word", "Create an editable Word document from typed Markdown.", ConversionInputKind.Text, ".md,.markdown,.txt", "MarkdownReader.Parse(markdown).ToWordDocument(options)", "ocx-route-card--word")
    ];

    public static ConversionRoute Default => All[0];

    public static ConversionRoute Find(string? id) =>
        All.FirstOrDefault(route => string.Equals(route.Id, id, StringComparison.OrdinalIgnoreCase)) ?? Default;
}
