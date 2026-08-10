namespace OfficeIMO.OpenDocument;

/// <summary>Kind of node in an ODT paragraph's ordered inline syntax.</summary>
public enum OdtInlineNodeKind {
    /// <summary>Plain text, including ODF space, tab, and line-break elements.</summary>
    Text,
    /// <summary>A styled text span.</summary>
    Span,
    /// <summary>A hyperlink.</summary>
    Hyperlink,
    /// <summary>An embedded image frame.</summary>
    Image,
    /// <summary>A collapsed bookmark marker.</summary>
    Bookmark,
    /// <summary>A bookmark range start marker.</summary>
    BookmarkStart,
    /// <summary>A bookmark range end marker.</summary>
    BookmarkEnd,
    /// <summary>An inline element not represented by the current typed surface.</summary>
    Other
}

/// <summary>
/// An ordered typed view of a direct child in an ODT paragraph. This syntax view keeps
/// mixed plain text, simple spans, simple hyperlinks, images, and bookmark markers in
/// document order. Nested inline markup is surfaced as <see cref="OdtInlineNodeKind.Other"/>
/// so converters cannot mistake a flattened representation for an exact mapping.
/// </summary>
public sealed class OdtInlineNode {
    private OdtInlineNode(OdtInlineNodeKind kind, string text, OdtSpan? span = null,
        OdtHyperlink? hyperlink = null, OdtImage? image = null, string? name = null,
        string? qualifiedName = null) {
        Kind = kind;
        Text = text;
        Span = span;
        Hyperlink = hyperlink;
        Image = image;
        Name = name;
        QualifiedName = qualifiedName;
    }

    /// <summary>Node kind.</summary>
    public OdtInlineNodeKind Kind { get; }
    /// <summary>Decoded text contributed by this node.</summary>
    public string Text { get; }
    /// <summary>Styled span for <see cref="OdtInlineNodeKind.Span"/>.</summary>
    public OdtSpan? Span { get; }
    /// <summary>Hyperlink for <see cref="OdtInlineNodeKind.Hyperlink"/>.</summary>
    public OdtHyperlink? Hyperlink { get; }
    /// <summary>Image for <see cref="OdtInlineNodeKind.Image"/>.</summary>
    public OdtImage? Image { get; }
    /// <summary>Bookmark name for bookmark marker nodes.</summary>
    public string? Name { get; }
    /// <summary>Expanded XML name for an unrepresented element.</summary>
    public string? QualifiedName { get; }

    internal static IReadOnlyList<OdtInlineNode> Read(
        OdtDocument document,
        XElement paragraph,
        string partPath) {
        // Enforce the paragraph-wide decoded-text budget before producing per-node values.
        _ = OdfTextCodec.Read(paragraph);
        var result = new List<OdtInlineNode>();
        var plainNodes = new List<XNode>();

        void FlushPlain() {
            if (plainNodes.Count == 0) return;
            string text = OdfTextCodec.ReadNodes(plainNodes);
            if (text.Length > 0) result.Add(new OdtInlineNode(OdtInlineNodeKind.Text, text));
            plainNodes.Clear();
        }

        foreach (XNode node in paragraph.Nodes()) {
            if (node is XText) {
                plainNodes.Add(node);
                continue;
            }
            if (!(node is XElement element)) continue;
            if (element.Name == OdfNamespaces.Text + "s"
                || element.Name == OdfNamespaces.Text + "tab"
                || element.Name == OdfNamespaces.Text + "line-break") {
                plainNodes.Add(element);
                continue;
            }

            FlushPlain();
            if ((element.Name == OdfNamespaces.Text + "span"
                    || element.Name == OdfNamespaces.Text + "a")
                && HasNestedInlineMarkup(element)) {
                result.Add(new OdtInlineNode(OdtInlineNodeKind.Other, OdfTextCodec.Read(element),
                    qualifiedName: element.Name.ToString()));
            } else if (element.Name == OdfNamespaces.Text + "span") {
                var span = new OdtSpan(document, element, partPath);
                result.Add(new OdtInlineNode(OdtInlineNodeKind.Span, span.Text, span: span));
            } else if (element.Name == OdfNamespaces.Text + "a") {
                var hyperlink = new OdtHyperlink(document, element, partPath);
                result.Add(new OdtInlineNode(OdtInlineNodeKind.Hyperlink, hyperlink.Text, hyperlink: hyperlink));
            } else if (element.Name == OdfNamespaces.Draw + "frame"
                && element.Element(OdfNamespaces.Draw + "image") != null) {
                var image = new OdtImage(document, element, partPath);
                result.Add(new OdtInlineNode(OdtInlineNodeKind.Image, string.Empty, image: image));
            } else if (element.Name == OdfNamespaces.Text + "bookmark") {
                result.Add(BookmarkNode(OdtInlineNodeKind.Bookmark, element));
            } else if (element.Name == OdfNamespaces.Text + "bookmark-start") {
                result.Add(BookmarkNode(OdtInlineNodeKind.BookmarkStart, element));
            } else if (element.Name == OdfNamespaces.Text + "bookmark-end") {
                result.Add(BookmarkNode(OdtInlineNodeKind.BookmarkEnd, element));
            } else {
                result.Add(new OdtInlineNode(OdtInlineNodeKind.Other, OdfTextCodec.Read(element),
                    qualifiedName: element.Name.ToString()));
            }
        }
        FlushPlain();
        return result;
    }

    private static OdtInlineNode BookmarkNode(OdtInlineNodeKind kind, XElement element) =>
        new OdtInlineNode(kind, string.Empty,
            name: (string?)element.Attribute(OdfNamespaces.Text + "name"));

    private static bool HasNestedInlineMarkup(XElement element) => element.Elements().Any(child =>
        child.Name != OdfNamespaces.Text + "s"
        && child.Name != OdfNamespaces.Text + "tab"
        && child.Name != OdfNamespaces.Text + "line-break");
}
