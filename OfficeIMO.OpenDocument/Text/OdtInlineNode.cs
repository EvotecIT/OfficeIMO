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
/// An ordered typed view of ODT inline syntax. Spans and hyperlinks retain their
/// child nodes in document order so nested formatting can be resolved by consumers.
/// </summary>
public sealed class OdtInlineNode {
    private OdtInlineNode(OdtInlineNodeKind kind, string text, OdtSpan? span = null,
        OdtHyperlink? hyperlink = null, OdtImage? image = null, string? name = null,
        string? qualifiedName = null, IReadOnlyList<OdtInlineNode>? children = null) {
        Kind = kind;
        Text = text;
        Span = span;
        Hyperlink = hyperlink;
        Image = image;
        Name = name;
        QualifiedName = qualifiedName;
        Children = children ?? Array.Empty<OdtInlineNode>();
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
    /// <summary>Ordered content inside a span or hyperlink; empty for leaf nodes.</summary>
    public IReadOnlyList<OdtInlineNode> Children { get; }

    internal static IReadOnlyList<OdtInlineNode> Read(
        OdtDocument document,
        XElement paragraph,
        string partPath) {
        // Enforce the paragraph-wide decoded-text budget before producing per-node values.
        _ = OdfTextCodec.Read(paragraph);
        return ReadChildren(document, paragraph, partPath);
    }

    private static IReadOnlyList<OdtInlineNode> ReadChildren(
        OdtDocument document, XElement parent, string partPath) {
        var result = new List<OdtInlineNode>();
        var plainNodes = new List<XNode>();

        void FlushPlain() {
            if (plainNodes.Count == 0) return;
            string text = OdfTextCodec.ReadNodes(plainNodes);
            if (text.Length > 0) result.Add(new OdtInlineNode(OdtInlineNodeKind.Text, text));
            plainNodes.Clear();
        }

        foreach (XNode node in parent.Nodes()) {
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
            if (element.Name == OdfNamespaces.Text + "span") {
                var span = new OdtSpan(document, element, partPath);
                result.Add(new OdtInlineNode(OdtInlineNodeKind.Span, span.Text, span: span,
                    children: ReadChildren(document, element, partPath)));
            } else if (element.Name == OdfNamespaces.Text + "a") {
                var hyperlink = new OdtHyperlink(document, element, partPath);
                result.Add(new OdtInlineNode(OdtInlineNodeKind.Hyperlink, hyperlink.Text, hyperlink: hyperlink,
                    children: ReadChildren(document, element, partPath)));
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

}
