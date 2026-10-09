using OfficeIMO.Html.Dom;
using System.Text.RegularExpressions;

namespace OfficeIMO.Chm;

public sealed partial class ChmDocument {
    private static readonly string[] ImageUrlAttributes = { "src", "data-src", "data-original", "data-original-src", "data-lazy-src" };
    private static readonly string[] ImageSetAttributes = { "srcset", "data-srcset", "data-original-srcset", "data-lazy-srcset" };
    private static readonly Regex LocalCssUrl = new Regex(@"\burl\(\s*(?<quote>['""]?)(?<reference>#[^ \t\r\n\f)'""]+)\k<quote>\s*\)",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(100));

    private void EmbedImages(HtmlElement element, ChmConversionOptions options, string sourcePath, ref long bytes,
        List<OfficeConversionFidelityDiagnostic> diagnostics, CancellationToken token) {
        foreach (string name in ImageUrlAttributes) {
            token.ThrowIfCancellationRequested();
            string? source = element.GetAttribute(name);
            if (source == null) continue;
            string? embedded = EmbedImageReference(source, options, sourcePath, ref bytes, diagnostics);
            if (embedded == null) element.RemoveAttribute(name);
            else element.SetAttribute(name, embedded);
        }
        foreach (string name in ImageSetAttributes) {
            string? source = element.GetAttribute(name);
            if (source == null) continue;
            var candidates = new List<string>();
            foreach (HtmlSrcSetCandidate candidate in HtmlSrcSetParser.Enumerate(source)) {
                token.ThrowIfCancellationRequested();
                string? embedded = EmbedImageReference(candidate.Url, options, sourcePath, ref bytes, diagnostics);
                if (embedded != null) candidates.Add(embedded + (candidate.Descriptor.Length == 0 ? string.Empty : " " + candidate.Descriptor));
            }
            if (candidates.Count == 0) element.RemoveAttribute(name);
            else element.SetAttribute(name, string.Join(", ", candidates));
        }
    }

    private string? EmbedImageReference(string source, ChmConversionOptions options, string sourcePath, ref long bytes,
        List<OfficeConversionFidelityDiagnostic> diagnostics) {
        ChmEntry? entry;
        if (Uri.TryCreate(source, UriKind.Absolute, out Uri? uri)) {
            if (uri.Scheme != "chm" || uri.Host != "archive") return source;
            entry = FindUriEntry(uri);
        } else entry = FindEntry(source, sourcePath);
        if (entry == null || entry.IsSystem || entry.IsDirectory) {
            diagnostics.Add(new OfficeConversionFidelityDiagnostic("CHM_IMAGE_MISSING", "An image reference has no embedded resource.", OfficeConversionLossKind.Omission, "OfficeIMO.Chm", sourcePath));
            return null;
        }
        if (entry.Length > options.MaxEmbeddedImageBytes || entry.Length > options.MaxTotalEmbeddedImageBytes - bytes)
            throw ChmBinary.Error("CONVERSION_LIMIT", "Embedded images exceed the configured projection budget.");
        bytes += entry.Length;
        return "data:" + GetMediaType(entry.Path) + ";base64," + Convert.ToBase64String(entry.GetBytes());
    }

    private void RewriteSvgReferences(HtmlElement element, string prefix, string sourcePath, Dictionary<string, string> anchors) {
        if (element.NamespaceUri != "http://www.w3.org/2000/svg") return;
        foreach (HtmlAttribute attribute in element.Attributes.ToArray()) {
            string? value = attribute.Value;
            if (attribute.LocalName == "href" && (attribute.NamespaceUri.Length == 0 || attribute.NamespaceUri == "http://www.w3.org/1999/xlink"))
                value = RewriteBookLink(value, sourcePath, anchors);
            else value = LocalCssUrl.Replace(value, match => {
                Group reference = match.Groups["reference"];
                int position = reference.Index - match.Index + 1;
                return match.Value.Insert(position, prefix + "-");
            });
            if (value == null) element.RemoveAttribute(attribute.NamespaceUri, attribute.LocalName);
            else if (value != attribute.Value) element.SetAttribute(attribute.Name, value, attribute.NamespaceUri);
        }
    }
}
