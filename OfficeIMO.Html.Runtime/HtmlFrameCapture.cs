using OfficeIMO.Html.Dom;

namespace OfficeIMO.Html.Runtime;

/// <summary>An inert child browsing-context snapshot attached to one iframe in its containing capture.</summary>
public sealed class HtmlFrameCapture {
    /// <summary>Creates an immutable frame snapshot from a frozen owned document.</summary>
    public HtmlFrameCapture(
        int frameElementNodeId,
        HtmlDocument document,
        Uri documentUrl,
        Uri baseUri,
        IReadOnlyList<HtmlFrameCapture>? frames = null) {
        if (frameElementNodeId <= 0) throw new ArgumentOutOfRangeException(nameof(frameElementNodeId));
        ArgumentNullException.ThrowIfNull(document);
        if (!document.IsReadOnly) throw new ArgumentException("A captured frame document must be frozen.", nameof(document));
        FrameElementNodeId = frameElementNodeId;
        Document = document;
        DocumentUrl = HtmlRuntimeResourcePolicy.ValidateUrl(documentUrl);
        BaseUri = ValidateBaseUri(baseUri);
        Frames = Array.AsReadOnly((frames ?? Array.Empty<HtmlFrameCapture>()).ToArray());
        ValidateChildren(document, Frames);
    }

    /// <summary>Node identity of the iframe in the containing document.</summary>
    public int FrameElementNodeId { get; }
    /// <summary>Frozen owned child document. Child scripts are not executed during capture or conversion.</summary>
    public HtmlDocument Document { get; }
    /// <summary>Final HTTP(S) identity of the child document.</summary>
    public Uri DocumentUrl { get; }
    /// <summary>Effective base URI frozen with the child document.</summary>
    public Uri BaseUri { get; }
    /// <summary>Nested same-origin frame snapshots in containing-document order.</summary>
    public IReadOnlyList<HtmlFrameCapture> Frames { get; }

    internal static void ValidateChildren(HtmlDocument containingDocument, IReadOnlyList<HtmlFrameCapture> frames) {
        var ids = new HashSet<int>();
        foreach (HtmlFrameCapture frame in frames) {
            if (frame == null || !ids.Add(frame.FrameElementNodeId))
                throw new ArgumentException("Captured frames must have unique non-null iframe identities.", nameof(frames));
            if (containingDocument.GetNode(frame.FrameElementNodeId) is not HtmlElement {
                    LocalName: "iframe", NamespaceUri: HtmlElement.HtmlNamespace
                }) {
                throw new ArgumentException("A captured frame identity must reference an iframe in its containing document.", nameof(frames));
            }
        }
    }

    internal static Uri ValidateBaseUri(Uri value) {
        ArgumentNullException.ThrowIfNull(value);
        if (!value.IsAbsoluteUri || value.AbsoluteUri.Length > 8192 || value.Scheme is "data" or "javascript")
            throw new ArgumentException("A captured frame base URI must be absolute and bounded.", nameof(value));
        return value;
    }
}
