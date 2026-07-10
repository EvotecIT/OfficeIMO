using System.Text;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

internal static class HtmlRenderStylesheetApplier {
    private const string ComponentName = "OfficeIMO.Html.Renderer";

    internal static void Apply(IHtmlDocument document, HtmlRenderResourceSet resources, HtmlDiagnosticReport diagnostics) {
        foreach (IElement link in document.QuerySelectorAll("link[href]")) {
            if (!IsStylesheetLink(link)) {
                continue;
            }

            string source = link.GetAttribute("href") ?? string.Empty;
            if (!resources.TryGet(source, null, out HtmlResolvedResource resource)) {
                continue;
            }

            if (!TryDecodeCss(resource.Bytes, out string css)) {
                diagnostics.Add(
                    ComponentName,
                    HtmlRenderDiagnosticCodes.StylesheetEncodingUnsupported,
                    "A resolved stylesheet could not be decoded as UTF-8 or BOM-declared UTF-16 CSS text.",
                    HtmlDiagnosticSeverity.Warning,
                    source,
                    resource.ContentType);
                continue;
            }

            if (HtmlResourcePipeline.HasNestedStylesheetResources(css)) {
                diagnostics.Add(
                    ComponentName,
                    HtmlRenderDiagnosticCodes.StylesheetNestedResourcesPending,
                    "The external stylesheet was applied, but its nested imports or URL resources require recursive stylesheet loading.",
                    HtmlDiagnosticSeverity.Warning,
                    source,
                    resource.ContentType);
            }

            IElement style = document.CreateElement("style");
            style.TextContent = css;
            style.SetAttribute("data-officeimo-source", source);
            string media = link.GetAttribute("media") ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(media)) {
                style.SetAttribute("media", media);
            }

            INode? parent = link.Parent;
            if (parent == null) {
                continue;
            }

            INode? next = link.NextSibling;
            if (next == null) {
                parent.AppendChild(style);
            } else {
                parent.InsertBefore(style, next);
            }
        }
    }

    private static bool IsStylesheetLink(IElement link) {
        string rel = link.GetAttribute("rel") ?? string.Empty;
        foreach (string token in rel.Split(new[] { ' ', '\t', '\r', '\n', '\f' }, StringSplitOptions.RemoveEmptyEntries)) {
            if (string.Equals(token, "stylesheet", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
        }

        return false;
    }

    private static bool TryDecodeCss(byte[] bytes, out string css) {
        css = string.Empty;
        try {
            if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) {
                css = new UTF8Encoding(false, true).GetString(bytes, 3, bytes.Length - 3);
            } else if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE) {
                css = new UnicodeEncoding(false, true, true).GetString(bytes, 2, bytes.Length - 2);
            } else if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF) {
                css = new UnicodeEncoding(true, true, true).GetString(bytes, 2, bytes.Length - 2);
            } else {
                css = new UTF8Encoding(false, true).GetString(bytes);
            }

            return true;
        } catch (DecoderFallbackException) {
            css = string.Empty;
            return false;
        }
    }
}
