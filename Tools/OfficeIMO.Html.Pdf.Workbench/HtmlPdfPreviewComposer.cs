using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;

namespace OfficeIMO.Html.Pdf.Workbench;

public static class HtmlPdfPreviewComposer {
    private const string Policy = "default-src 'none'; img-src data: blob:; style-src 'unsafe-inline' data: blob:; font-src data: blob:; media-src data: blob:; connect-src 'none'; frame-src 'none'; object-src 'none'; base-uri 'none'; form-action 'none'";

    public static string Compose(string html, string css) => ComposeCore(html, css, preview: true);

    public static string ComposeForCapture(string html, string css, string? language = null) =>
        ComposeCore(html, css, preview: false, language);

    private static string ComposeCore(string html, string css, bool preview, string? language = null) {
        var parser = new HtmlParser();
        IHtmlDocument document = parser.ParseDocument(html ?? string.Empty);
        IElement head = document.Head ?? throw new InvalidOperationException("The HTML parser did not create a head element.");

        foreach (IElement meta in document.QuerySelectorAll("meta[http-equiv]").ToArray()) {
            string directive = meta.GetAttribute("http-equiv") ?? string.Empty;
            if (preview ||
                string.Equals(directive, "Content-Security-Policy", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(directive, "Refresh", StringComparison.OrdinalIgnoreCase)) {
                meta.Remove();
            }
        }

        if (!string.IsNullOrWhiteSpace(language)) {
            document.DocumentElement.SetAttribute("lang", language.Trim());
        }

        if (preview) {
            foreach (IElement activeContainer in document.QuerySelectorAll("base,iframe,frame,object,embed").ToArray()) activeContainer.Remove();

            IElement policy = document.CreateElement("meta");
            policy.SetAttribute("http-equiv", "Content-Security-Policy");
            policy.SetAttribute("content", Policy);
            head.Prepend(policy);
        }

        IElement style = document.CreateElement("style");
        style.TextContent = EscapeStyleTerminator(css ?? string.Empty);
        head.Append(style);
        return "<!doctype html>" + document.DocumentElement.OuterHtml;
    }

    private static string EscapeStyleTerminator(string css) =>
        css.Replace("</style", "<\\/style", StringComparison.OrdinalIgnoreCase);
}
