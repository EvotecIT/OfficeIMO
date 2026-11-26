using OfficeIMO.Drawing;
using System.Xml.Linq;
using OfficeIMO.Epub;

namespace OfficeIMO.Fb2;

public static partial class Fb2Reader {
    private sealed partial class Projection {
        private readonly Dictionary<string, EpubManifestItem> _images = new(StringComparer.Ordinal);
        private readonly Dictionary<string, (int Body, string Id)> _targets = new(StringComparer.Ordinal);
        private readonly Dictionary<XElement, string> _elementIds = new();
        private int _anchorCount;

        private void AddImages(EpubPublication book, XElement root) {
            int index = 0;
            foreach (XElement binary in root.Elements(Fb + "binary")) {
                token.ThrowIfCancellationRequested();
                if (++index > 1_000) throw new InvalidDataException("FictionBook exceeds 1,000 embedded binaries.");
                string? id = (string?)binary.Attribute("id");
                string? type = (string?)binary.Attribute("content-type");
                string? extension = type switch {
                    "image/jpeg" => ".jpg",
                    "image/png" => ".png",
                    "image/gif" => ".gif",
                    "image/webp" => ".webp",
                    _ => null
                };
                if (string.IsNullOrEmpty(id) || _images.ContainsKey(id!) || extension == null) {
                    Report("FB2_BINARY_OMITTED", "An embedded binary is missing an identifier, uses a duplicate identifier, or has an unsupported media type.", id);
                    continue;
                }
                byte[] data;
                try { data = Convert.FromBase64String(binary.Value); }
                catch (FormatException) {
                    Report("FB2_BINARY_INVALID", "An embedded image has invalid base64 data.", id);
                    continue;
                }
                if (!EpubImportedPublication.TryReadRaster(data, null, token, out OfficeImageInfo actual) ||
                    !string.Equals(type, actual.MimeType, StringComparison.OrdinalIgnoreCase)) {
                    Report("FB2_BINARY_INVALID", "An embedded image does not match its declared media type.", id);
                    continue;
                }
                string asset = "image-" + Ordinal(index);
                _images.Add(id!, book.AddResource(asset, "EPUB/" + asset + extension, type!, data));
            }
        }

        private static string BodyName(int body) => "body-" + Ordinal(body + 1) + ".xhtml";

        private void IndexAnchors(XElement body, int number) {
            foreach (XElement element in body.DescendantsAndSelf()) {
                token.ThrowIfCancellationRequested();
                string? id = (string?)element.Attribute("id");
                if (id == null && element.Name != Fb + "section") continue;
                string anchor = "fb2-" + Ordinal(++_anchorCount);
                _elementIds.Add(element, anchor);
                if (id != null && _targets.ContainsKey(id))
                    Report("FB2_DUPLICATE_ANCHOR", "A duplicate FictionBook identifier resolves to its first occurrence.", id);
                else if (id != null)
                    _targets.Add(id, (number, anchor));
            }
        }

        private IEnumerable<EpubNavigationEntry> ReadNavigation(XElement parent, string path) {
            foreach (XElement section in parent.Elements(Fb + "section")) {
                token.ThrowIfCancellationRequested();
                var children = ReadNavigation(section, path).ToArray();
                string? label = Value(section, "title");
                if (label == null) {
                    foreach (EpubNavigationEntry child in children) yield return child;
                } else {
                    yield return new EpubNavigationEntry(label, path + "#" + _elementIds[section], children);
                }
            }
        }

        private object? Render(XNode node, int body, int level) {
            token.ThrowIfCancellationRequested();
            if (node is XText text) return new XText(text.Value);
            if (node is not XElement source) return null;
            string name = source.Name.LocalName;
            string? tag = source.Name.Namespace == Fb ? name switch {
                "section" => "section",
                "title" => "h" + Ordinal(Math.Min(level, 6)),
                "p" => source.Parent?.Name == Fb + "title" ? "span" : "p",
                "subtitle" => "h" + Ordinal(Math.Min(level + 1, 6)),
                "emphasis" => "em", "strong" => "strong", "strikethrough" => "s",
                "sub" => "sub", "sup" => "sup", "code" => "code", "style" => "span",
                "poem" or "stanza" or "v" => "div", "cite" or "epigraph" => "blockquote",
                "text-author" or "date" => "p", "empty-line" => "br",
                "table" => "table", "tr" => "tr", "td" => "td", "th" => "th",
                "a" => "a", "image" => "img", "annotation" => "aside",
                _ => null
            } : null;
            if (tag == null) {
                Report("FB2_ELEMENT_APPROXIMATED", "The source element's text was retained without its original structure.", name);
                tag = "span";
            }
            var target = new XElement(tag);
            if (_elementIds.TryGetValue(source, out string? anchor)) target.SetAttributeValue("id", anchor);
            if (source.Attribute(XNamespace.Xml + "lang") is { } lang) target.Add(new XAttribute(lang));
            if (tag == "a") {
                if ((string?)source.Attribute(Link + "href") is { } href) {
                    if (href.StartsWith("#", StringComparison.Ordinal) && _targets.TryGetValue(href.Substring(1), out var entry))
                        target.SetAttributeValue("href", (entry.Body == body ? "" : BodyName(entry.Body)) + "#" + entry.Id);
                    else if (EpubImportedPublication.IsWebLink(href))
                        target.SetAttributeValue("href", href);
                    else Report("FB2_LINK_OMITTED", "A link target could not be resolved safely.", href);
                }
            } else if (tag == "img") {
                string? href = (string?)source.Attribute(Link + "href");
                string alt = (string?)source.Attribute("alt") ?? (string?)source.Attribute("title") ?? "";
                if (href is { Length: > 1 } && href[0] == '#' && _images.TryGetValue(href.Substring(1), out EpubManifestItem? image)) {
                    target.SetAttributeValue("src", image.Href);
                    target.SetAttributeValue("alt", alt);
                    return target;
                }
                Report("FB2_IMAGE_UNAVAILABLE", "An image reference does not resolve to a supported embedded raster image.", href);
                return new XElement("span", target.Attributes(), new XText(alt));
            }
            if (tag is "td" or "th") {
                foreach (string span in new[] { "colspan", "rowspan" })
                    if (int.TryParse((string?)source.Attribute(span), out int count) && count is > 0 and <= 100)
                        target.SetAttributeValue(span, Ordinal(count));
            }
            if (name == "style") Report("FB2_STYLE_OMITTED", "FictionBook named styles are omitted. Semantic formatting is retained.");
            if (name == "title") {
                bool previousParagraph = false;
                foreach (XNode child in source.Nodes()) {
                    bool paragraph = child is XElement item && item.Name == Fb + "p";
                    if (previousParagraph && paragraph) target.Add(new XElement("br"));
                    target.Add(Render(child, body, level));
                    previousParagraph |= paragraph;
                }
            } else if (tag != "br") {
                target.Add(source.Nodes().Select(child => Render(child, body, level + (name == "section" ? 1 : 0))));
            }
            return target;
        }
    }
}
