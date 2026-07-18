using OfficeIMO.Drawing;
using OfficeIMO.OneNote.Markdown;
using System.Collections.Generic;
using System.Net;

namespace OfficeIMO.OneNote.Html;

internal static class OneNoteVisualHtmlRenderer {
    private const string DefaultStyles = ".officeimo-onenote-visual{margin:0 auto;max-width:max-content}.officeimo-onenote-page{margin:0 0 1.5rem;box-shadow:0 .2rem .8rem rgba(0,0,0,.14);background:#fff}.officeimo-onenote-page svg{display:block;max-width:100%;height:auto}.officeimo-onenote-assistive{position:absolute!important;width:1px!important;height:1px!important;padding:0!important;margin:-1px!important;overflow:hidden!important;clip:rect(0,0,0,0)!important;white-space:pre-wrap!important;border:0!important}";

    internal static string RenderDocument(string title, IReadOnlyList<OneNotePageReference> pages, OneNoteVisualHtmlOptions? options) {
        OneNoteVisualHtmlOptions effective = Prepare(options);
        string documentTitle = string.IsNullOrWhiteSpace(effective.DocumentTitle) ? title : effective.DocumentTitle!;
        var html = new StringBuilder(4096);
        html.Append("<!DOCTYPE html><html lang=\"").Append(Attribute(effective.Language)).Append("\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\"><title>")
            .Append(Text(documentTitle)).Append("</title>");
        if (effective.IncludeDefaultStyles) html.Append("<style>").Append(DefaultStyles).Append("</style>");
        html.Append("</head><body>").Append(RenderFragmentCore(pages, effective)).Append("</body></html>");
        return html.ToString();
    }

    internal static string RenderFragment(IReadOnlyList<OneNotePageReference> pages, OneNoteVisualHtmlOptions? options) =>
        RenderFragmentCore(pages, Prepare(options));

    private static string RenderFragmentCore(IReadOnlyList<OneNotePageReference> pages, OneNoteVisualHtmlOptions options) {
        var html = new StringBuilder(4096);
        html.Append("<main class=\"officeimo-onenote-visual\">");
        foreach (OneNotePageReference reference in pages) {
            string id = "officeimo-onenote-page-" + reference.Index.ToString(System.Globalization.CultureInfo.InvariantCulture);
            string title = string.IsNullOrWhiteSpace(reference.Page.Title) ? "Untitled page" : reference.Page.Title;
            OneNotePageVisualSnapshot snapshot = OneNotePageRenderer.CreateSnapshot(reference.Page, options.PageRendering);
            var diagnostics = new List<OfficeImageExportDiagnostic>(snapshot.Diagnostics);
            var fallbackCodec = new OfficeRasterImageFallbackCodec(
                options.PageRendering.ImageCodec,
                diagnostics,
                reference.SectionPath + "/" + title);
            string svg = OfficeDrawingSvgExporter.ToSvg(
                snapshot.Drawing,
                options.PageRendering.Scale,
                OfficeSvgSizeUnit.Pixel,
                fallbackCodec,
                id + "-");
            if (options.DiagnosticSink != null) {
                foreach (OfficeImageExportDiagnostic diagnostic in diagnostics) options.DiagnosticSink.Add(diagnostic);
            }
            html.Append("<figure class=\"officeimo-onenote-page\" data-page-index=\"")
                .Append(reference.Index.ToString(System.Globalization.CultureInfo.InvariantCulture))
                .Append("\" data-section-path=\"").Append(Attribute(reference.SectionPath))
                .Append("\" aria-labelledby=\"").Append(id).Append("\">")
                .Append(svg)
                .Append("<figcaption id=\"").Append(id).Append("\" class=\"officeimo-onenote-assistive\">")
                .Append(Text(title));
            if (options.IncludeAccessibleText) {
                string semanticText = OneNoteMarkdownProjection.ToText(reference.Page);
                if (!string.IsNullOrWhiteSpace(semanticText)) html.Append("\n").Append(Text(semanticText));
            }
            html.Append("</figcaption></figure>");
        }
        html.Append("</main>");
        return html.ToString();
    }

    private static OneNoteVisualHtmlOptions Prepare(OneNoteVisualHtmlOptions? options) {
        OneNoteVisualHtmlOptions effective = options?.Clone() ?? new OneNoteVisualHtmlOptions();
        effective.Validate();
        return effective;
    }

    private static string Text(string? value) => WebUtility.HtmlEncode(value ?? string.Empty);
    private static string Attribute(string? value) => WebUtility.HtmlEncode(value ?? string.Empty);
}
