using OfficeIMO.Html.Dom;

namespace OfficeIMO.Html.Runtime;

// Flat transport preserves authored DOM structure without serializing and reparsing HTML.
// Parent indices always precede children, so materialization and depth validation are iterative.
internal sealed class HtmlRuntimeWireDocument {
    public string ProviderId { get; set; } = string.Empty;
    public HtmlDocumentMode Mode { get; set; }
    public string? Error { get; set; }
    public List<HtmlRuntimeWireNode> Nodes { get; set; } = new();
    public Uri DocumentUrl { get; set; } = new("https://officeimo.invalid/");
    public Uri? BaseUri { get; set; }
    public List<HtmlRuntimeResource> Resources { get; set; } = new();
    public List<HtmlRuntimeWireFrame> Frames { get; set; } = new();

    internal HtmlDocument Materialize(IHtmlDomServices services, HtmlScriptRequest request, CancellationToken token, out IReadOnlyList<HtmlFrameCapture> frames) {
        int nodeCount = 0;
        HtmlDocument document = MaterializeDocument(this, services, request, token, ref nodeCount, out frames);
        return document;
    }

    private static HtmlDocument MaterializeDocument(
        HtmlRuntimeWireDocument wire,
        IHtmlDomServices services,
        HtmlScriptRequest request,
        CancellationToken token,
        ref int nodeCount,
        out IReadOnlyList<HtmlFrameCapture> frames) {
        if (wire.Nodes == null || wire.Frames == null || wire.Nodes.Count > request.MaxNodes - nodeCount)
            throw new HtmlScriptRuntimeException("Captured node budget exceeded.");
        nodeCount += wire.Nodes.Count;
        var document = new HtmlDocument(services, wire.ProviderId, wire.Mode);
        var nodes = new List<HtmlNode> { document };
        var depths = new List<int> { 0 };
        foreach (HtmlRuntimeWireNode entry in wire.Nodes) {
            token.ThrowIfCancellationRequested();
            if (entry == null || entry.Parent < 0 || entry.Parent >= nodes.Count) throw new HtmlScriptRuntimeException("Invalid capture parent.");
            HtmlNode parent = nodes[entry.Parent];
            int depth = depths[entry.Parent] + (entry.Kind == HtmlNodeKind.Element ? 1 : 0);
            if (depth > request.MaxDepth) throw new HtmlScriptRuntimeException("Captured depth budget exceeded.");
            HtmlNode node;
            if (entry.IsTemplateContent) {
                if (parent is not HtmlElement template || template.TemplateContent != null || entry.Kind != HtmlNodeKind.DocumentFragment)
                    throw new HtmlScriptRuntimeException("Invalid template capture.");
                node = document.GetOrCreateTemplateContent(template);
            } else {
                node = entry.Kind switch {
                    HtmlNodeKind.Element => document.CreateElement(entry.Name, entry.NamespaceUri, entry.Prefix),
                    HtmlNodeKind.Text => document.CreateTextNode(entry.Data),
                    HtmlNodeKind.Comment => document.CreateComment(entry.Data),
                    HtmlNodeKind.DocumentType => document.CreateDocumentType(entry.Name, entry.PublicIdentifier, entry.SystemIdentifier),
                    _ => throw new HtmlScriptRuntimeException("Unsupported captured node kind.")
                };
                parent.AppendChild(node);
            }
            if (node is HtmlElement element) {
                foreach (HtmlRuntimeWireAttribute attribute in entry.Attributes) {
                    token.ThrowIfCancellationRequested();
                    element.SetAttribute(new HtmlAttribute(attribute.Name, attribute.Value, attribute.NamespaceUri));
                }
                element.FormState = entry.FormState;
            }
            nodes.Add(node);
            depths.Add(depth);
        }
        token.ThrowIfCancellationRequested();
        document.Freeze();
        var materializedFrames = new List<HtmlFrameCapture>(wire.Frames.Count);
        foreach (HtmlRuntimeWireFrame wireFrame in wire.Frames) {
            token.ThrowIfCancellationRequested();
            if (wireFrame == null || wireFrame.Document == null)
                throw new HtmlScriptRuntimeException("Invalid captured frame.");
            HtmlDocument child = MaterializeDocument(wireFrame.Document, services, request, token, ref nodeCount, out IReadOnlyList<HtmlFrameCapture> children);
            materializedFrames.Add(new HtmlFrameCapture(
                wireFrame.FrameElementNodeId,
                child,
                wireFrame.Document.DocumentUrl,
                wireFrame.Document.BaseUri ?? wireFrame.Document.DocumentUrl,
                children));
        }
        frames = Array.AsReadOnly(materializedFrames.ToArray());
        HtmlFrameCapture.ValidateChildren(document, frames);
        return document;
    }
}

internal sealed class HtmlRuntimeWireFrame {
    public int FrameElementNodeId { get; set; }
    public HtmlRuntimeWireDocument? Document { get; set; }
}

internal sealed class HtmlRuntimeWireNode {
    public int Parent { get; set; }
    public HtmlNodeKind Kind { get; set; }
    public string Name { get; set; } = string.Empty;
    public string NamespaceUri { get; set; } = string.Empty;
    public string? Prefix { get; set; }
    public string Data { get; set; } = string.Empty;
    public string PublicIdentifier { get; set; } = string.Empty;
    public string SystemIdentifier { get; set; } = string.Empty;
    public bool IsTemplateContent { get; set; }
    public List<HtmlRuntimeWireAttribute> Attributes { get; set; } = new();
    public HtmlFormControlState? FormState { get; set; }
}

internal sealed class HtmlRuntimeWireAttribute {
    public string Name { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
    public string NamespaceUri { get; set; } = string.Empty;
}
