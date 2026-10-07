namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private static readonly XNamespace OverlaySvg = "http://www.w3.org/2000/svg";
    private static readonly HashSet<string> NarratableSvgElements = new HashSet<string>(new[] {
        "svg", "g", "a", "switch", "text", "tspan", "textPath", "image", "use", "path", "rect", "circle", "ellipse", "line", "polyline", "polygon"
    }, StringComparer.Ordinal);
    private static readonly HashSet<string> NonNarratableSvgContainers = new HashSet<string>(new[] {
        "defs", "symbol", "clipPath", "mask", "marker", "pattern", "linearGradient", "radialGradient", "filter", "metadata", "title", "desc", "foreignObject"
    }, StringComparer.Ordinal);

    private XDocument EditableOverlayContent(EpubManifestItem item) {
        if (!IsSupportedContentDocument(item.MediaType)) throw new NotSupportedException("Narration authoring requires XHTML or SVG content.");
        string path = RequireLocalPath(item);
        if (_encryption.Any(entry => entry.Path == path)) throw new NotSupportedException("Encrypted content cannot be narrated.");
        XDocument document = GetContentXml(item.Id);
        ValidateContent(document, item.MediaType);
        return document;
    }

    // This is a structural target check. CSS visibility, switch selection, use instances and
    // reading-system highlighting need rendered/native qualification.
    private static bool IsNarratableElement(XElement element) {
        if (element.AncestorsAndSelf().Any(parent => parent.Name == Html + "template" ||
            parent.Name.Namespace == OverlaySvg && NonNarratableSvgContainers.Contains(parent.Name.LocalName))) return false;
        if (element.Name.Namespace == Html)
            return !new[] { "script", "style", "link", "meta", "template" }.Contains(element.Name.LocalName);
        return element.Name.Namespace == OverlaySvg && NarratableSvgElements.Contains(element.Name.LocalName);
    }
}
