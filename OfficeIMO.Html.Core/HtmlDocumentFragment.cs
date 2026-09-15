namespace OfficeIMO.Html.Dom;

/// <summary>
/// An owned document fragment. A fragment can retain multiple top-level nodes and splices its
/// children when appended to an element or document.
/// </summary>
public sealed class HtmlDocumentFragment : HtmlNode {
    internal HtmlDocumentFragment(HtmlDocument document, int nodeId)
        : base(document, nodeId, HtmlNodeKind.DocumentFragment) {
    }
}
