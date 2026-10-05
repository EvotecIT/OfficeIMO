using OfficeIMO.Html;
using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private static bool RewriteMovedXml(XDocument document, string owner, string destination, string oldPath, string newPath, CancellationToken token, ContentReferenceMap? map = null) {
        if (document.Root == null) throw new InvalidDataException("Resource XML has no root.");
        if (document.Descendants().Attributes(XNamespace.Xml + "base").Any())
            throw new NotSupportedException("Resource renaming does not support XML base declarations.");
        if (document.Descendants().Any(element => element.Name.LocalName == "script" || element.Attributes().Any(attribute =>
            attribute.Name.NamespaceName.Length == 0 && (attribute.Name.LocalName.StartsWith("on", StringComparison.OrdinalIgnoreCase) || attribute.Name.LocalName is "srcdoc" or "codebase"))))
            throw new NotSupportedException("Resource renaming requires non-scripted content without embedded documents or code bases.");
        if (document.Descendants().Any(element =>
            (element.Name.NamespaceName == "http://www.w3.org/2000/svg" && new[] { "animate", "set", "animateMotion", "animateTransform" }.Contains(element.Name.LocalName, StringComparer.Ordinal)) ||
            (element.Name == Html + "meta" && string.Equals((string?)element.Attribute("http-equiv"), "refresh", StringComparison.OrdinalIgnoreCase))))
            throw new NotSupportedException("Resource renaming requires static content without SVG animation or refresh navigation.");
        XAttribute? baseAttribute = document.Root.Name == Html + "html" ? document.Root.Element(Html + "head")?.Elements(Html + "base")
            .Attributes("href").FirstOrDefault() : null;
        string? oldBase = baseAttribute?.Value;
        bool changed = RewriteMovedXmlStylesheets(document, owner, destination, oldPath, newPath, token, map);
        if (!string.IsNullOrWhiteSpace(oldBase)) {
            EpubReference reference = EpubReference.Resolve(owner, oldBase!, "#");
            if (!reference.IsValid) throw new InvalidDataException("Invalid content base URL.");
            if (reference.Kind == EpubReferenceKind.Container) {
                string target = reference.ContainerPath == oldPath ? newPath : reference.ContainerPath!;
                EpubReference current = EpubReference.Resolve(destination, oldBase!, "#");
                if (current.ContainerPath != target) {
                    baseAttribute!.Value = RelativeHref(destination, target) + (reference.Query == null ? string.Empty : "?" + reference.Query);
                    changed = true;
                }
            }
        }
        string? newBase = baseAttribute?.Value;
        string Rewrite(string value) => RewriteMovedReference(owner, oldBase, destination, newBase, value, oldPath, newPath, map);
        foreach (XElement element in document.Descendants()) {
            token.ThrowIfCancellationRequested();
            string ns = element.Name.NamespaceName;
            bool html = element.Name.Namespace == Html;
            bool svg = ns == "http://www.w3.org/2000/svg";
            bool math = ns == "http://www.w3.org/1998/Math/MathML";
            bool smil = ns == "http://www.w3.org/ns/SMIL";
            foreach (XAttribute attribute in element.Attributes().ToArray()) {
                if (element.Name == Html + "base" || attribute.IsNamespaceDeclaration) continue;
                // HTML image maps use a document-local name, independent of the document's base URL.
                if (html && attribute.Name == "usemap" && attribute.Value.StartsWith("#", StringComparison.Ordinal)) continue;
                string? replacement = null;
                bool uri = attribute.Name == "href" || attribute.Name == "src" || attribute.Name == "poster" ||
                    attribute.Name == XName.Get("href", "http://www.w3.org/1999/xlink") ||
                    (html && (attribute.Name == "cite" || attribute.Name == "longdesc" || attribute.Name == "usemap" ||
                        (element.Name == Html + "object" && attribute.Name == "data"))) ||
                    (math && attribute.Name == "definitionURL") || (smil && attribute.Name == Ops + "textref");
                if (uri) {
                    if (!html && !svg && !math && !smil && element.Name != Ncx + "content")
                        throw new NotSupportedException("Cannot rewrite a reference in an unknown XML vocabulary: " + element.Name);
                    bool documentLink = ((html || svg) && element.Name.LocalName == "a" || html && element.Name.LocalName == "area") &&
                        (attribute.Name == "href" || attribute.Name == XName.Get("href", "http://www.w3.org/1999/xlink"));
                    replacement = RewriteMovedReference(owner, oldBase, destination, newBase, attribute.Value, oldPath, newPath, map,
                        allowEmptyDocumentLink: documentLink);
                } else if (html && (attribute.Name == "srcset" || attribute.Name == "imagesrcset")) {
                    var result = new StringBuilder(attribute.Value);
                    foreach (var candidate in HtmlSrcSetParser.Enumerate(attribute.Value).Reverse()) {
                        token.ThrowIfCancellationRequested();
                        string url = Rewrite(candidate.Url);
                        if (url != candidate.Url) result.Remove(candidate.UrlStart, candidate.Url.Length).Insert(candidate.UrlStart, url);
                    }
                    replacement = result.ToString();
                } else if ((html || svg || math) && attribute.Name == "style") {
                    replacement = RewriteMovedCss(attribute.Value, Rewrite, includeFragmentReferences: map != null);
                } else if (svg && attribute.Name.NamespaceName.Length == 0 && new[] {
                    "fill", "stroke", "filter", "clip-path", "mask", "marker", "marker-start", "marker-mid", "marker-end", "cursor"
                }.Contains(attribute.Name.LocalName, StringComparer.Ordinal)) {
                    replacement = RewriteMovedCss(attribute.Value, Rewrite, includeFragmentReferences: map != null);
                } else if (html && (attribute.Name == "ping" || (element.Name == Html + "object" && attribute.Name == "archive"))) {
                    replacement = string.Join(" ", Tokens(attribute.Value).Select(Rewrite));
                }
                if (replacement != null && replacement != attribute.Value) { attribute.Value = replacement; changed = true; }
            }
            if ((html || svg) && element.Name.LocalName == "style") {
                string css = RewriteMovedCss(element.Value, Rewrite, includeFragmentReferences: map != null);
                if (css != element.Value) { element.Value = css; changed = true; }
            }
        }
        return changed;
    }
}
