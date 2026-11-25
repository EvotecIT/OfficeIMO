using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Dom;

namespace OfficeIMO.Notebook.Import;

internal static partial class RNotebookHtmlImporter {
    private const int MaxImageBytes = 5 * 1024 * 1024;
    private const int MaxTableCharacters = 2 * 1024 * 1024;

    private static NotebookOutput? ReadImage(HtmlElement element) {
        const string png = "data:image/png;base64,";
        const string jpeg = "data:image/jpeg;base64,";
        string? source = element.GetAttribute("src");
        if (source == null) return null;
        string mime;
        string encoded;
        if (source.StartsWith(png, StringComparison.OrdinalIgnoreCase)) {
            mime = "image/png";
            encoded = source.Substring(png.Length);
        } else if (source.StartsWith(jpeg, StringComparison.OrdinalIgnoreCase)) {
            mime = "image/jpeg";
            encoded = source.Substring(jpeg.Length);
        } else {
            return null;
        }
        if (encoded.Length > (MaxImageBytes + 2L) / 3L * 4L) return null;
        try {
            byte[] bytes = Convert.FromBase64String(encoded);
            if (bytes.Length > MaxImageBytes
                || !OfficeImageReader.TryIdentifyByContent(bytes, null, out OfficeImageInfo image)
                || image.Width < 1 || image.Height < 1 || (long)image.Width * image.Height > 16_000_000
                || !string.Equals(image.MimeType, mime, StringComparison.Ordinal)) return null;
            return Display(mime, encoded);
        } catch (FormatException) {
            return null;
        }
    }

    private static NotebookOutput? ReadTable(HtmlElement element, CancellationToken token, out TableOmissions omissions) {
        omissions = TableOmissions.None;
        string html = element.OuterHtml;
        if (html.Length > MaxTableCharacters) return null;
        HtmlUrlPolicy urlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile();
        TableOmissions noticed = InspectTableElement(element, urlPolicy);
        foreach (HtmlNode node in element.Descendants()) {
            token.ThrowIfCancellationRequested();
            if (node is HtmlElement descendant) noticed |= InspectTableElement(descendant, urlPolicy);
        }

        HtmlUrlPolicy resourcePolicy = HtmlUrlPolicy.CreateEmbeddedResourceProfile();
        resourcePolicy.ResolvedUrlTransform = _ => {
            noticed |= TableOmissions.Resource;
            return null;
        };
        var options = new HtmlNormalizationOptions {
            UrlPolicy = urlPolicy,
            ResourceUrlPolicy = resourcePolicy,
            PreserveStyleElements = false,
            Limits = HtmlConversionLimits.UntrustedDefaults
        };
        options.Limits.MaxInputCharacters = MaxTableCharacters;
        options.Limits.MaxHtmlNodes = 30_000;
        options.Limits.MaxHtmlDepth = 64;
        try {
            string normalized = HtmlNormalizer.Normalize(html, options);
            token.ThrowIfCancellationRequested();
            omissions = noticed;
            return normalized.Length > 0 ? Display("text/html", normalized) : null;
        } catch (HtmlDomLimitException) {
            return null;
        }
    }

    private static TableOmissions InspectTableElement(HtmlElement element, HtmlUrlPolicy urlPolicy) {
        // Normalization intentionally removes resources, event handlers and scripts before the
        // ordinary notebook projector receives this table. Observe them while the original
        // subtree is still available, including URLs rejected before the transform runs.
        TableOmissions noticed = IsUnsupportedSavedOutputElement(element.LocalName)
            ? TableOmissions.Structure : TableOmissions.None;
        foreach (HtmlAttribute attribute in element.Attributes) {
            if (attribute.Name.StartsWith("on", StringComparison.OrdinalIgnoreCase)) {
                noticed |= TableOmissions.Structure;
            } else if (string.IsNullOrWhiteSpace(attribute.Value)) {
                continue;
            } else if (attribute.Name is "src" or "srcset" or "srcdoc" or "background" or "poster") {
                noticed |= TableOmissions.Resource;
            } else if (attribute.Name == "href" && (element.LocalName is "a" or "area")
                && HtmlUrlPolicyEvaluator.ResolveUrl(attribute.Value, null, urlPolicy).Length == 0) {
                noticed |= TableOmissions.Structure;
            }
        }
        return noticed;
    }

    private static bool IsUnsupportedSavedOutputElement(string name) => IsIgnoredSavedMarkup(name)
        || name is "svg" or "canvas" or "video" or "audio" or "source"
            or "form" or "input" or "button" or "select" or "textarea";

    private static bool IsIgnoredSavedMarkup(string name) =>
        name is "script" or "style" or "template" or "iframe" or "object" or "embed";

    [Flags]
    private enum TableOmissions {
        None = 0,
        Resource = 1,
        Structure = 2
    }

    private static NotebookOutput Display(string mime, string value) =>
        NotebookOutput.CreateDisplayData(new Dictionary<string, JsonElement>(StringComparer.Ordinal) {
            [mime] = NotebookImportContext.JsonString(value)
        });
}
