using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Html.Dom;

/// <summary>An owned HTML tree with explicit snapshot and mutation semantics. Mutable documents require a single owner.</summary>
public sealed class HtmlDocument : HtmlNode {
    private int _nextNodeId;
    private readonly Dictionary<int, HtmlNode> _nodes = new Dictionary<int, HtmlNode>();

    /// <summary>Creates an empty mutable document for a parser or programmatic authoring.</summary>
    public HtmlDocument(IHtmlDomServices services, string providerId, HtmlDocumentMode mode = HtmlDocumentMode.Standards) : base(null, 0, HtmlNodeKind.Document) {
        Services = services ?? throw new ArgumentNullException(nameof(services));
        ProviderId = providerId ?? throw new ArgumentNullException(nameof(providerId));
        if (!Enum.IsDefined(typeof(HtmlDocumentMode), mode)) throw new ArgumentOutOfRangeException(nameof(mode));
        Mode = mode;
        _nodes.Add(0, this);
    }
    /// <summary>Unique identity of this tree instance; clones have independent identities.</summary>
    public Guid SnapshotId { get; } = Guid.NewGuid();
    /// <summary>Mutation revision used to invalidate derived state.</summary>
    public long Revision { get; private set; }
    /// <summary>Whether this tree rejects further mutation, including detached nodes.</summary>
    public bool IsReadOnly { get; private set; }
    /// <summary>Parser identity attached when the tree was created.</summary>
    public string ProviderId { get; }
    /// <summary>Layout compatibility mode established by the parser.</summary>
    public HtmlDocumentMode Mode { get; }
    /// <summary>Syntax services used for queries and serialization.</summary>
    public IHtmlDomServices Services { get; }
    /// <summary>Root element, if present.</summary>
    public HtmlElement? DocumentElement => Children.FirstOrDefault();
    /// <summary>Direct HTML body or frameset child of the root.</summary>
    public HtmlElement? Body => DocumentElement?.Children.FirstOrDefault(element => element.NamespaceUri == HtmlElement.HtmlNamespace && (element.LocalName == "body" || element.LocalName == "frameset"));
    /// <summary>Direct HTML head child of the root.</summary>
    public HtmlElement? Head => DocumentElement?.Children.FirstOrDefault(element => element.NamespaceUri == HtmlElement.HtmlNamespace && element.LocalName == "head");

    /// <summary>Creates a detached element in this mutable document.</summary>
    public HtmlElement CreateElement(string localName, string namespaceUri = HtmlElement.HtmlNamespace, string? prefix = null) {
        EnsureMutable();
        if (string.IsNullOrWhiteSpace(localName)) throw new ArgumentException("An element name is required.", nameof(localName));
        if (namespaceUri == null) throw new ArgumentNullException(nameof(namespaceUri));
        if (namespaceUri == HtmlElement.HtmlNamespace) localName = HtmlNames.LowerAscii(localName);
        return Register(new HtmlElement(this, NextId(), localName, namespaceUri, prefix));
    }
    /// <summary>Creates decoded text in this mutable document.</summary>
    public HtmlNode CreateTextNode(string text) => CreateDataNode(HtmlNodeKind.Text, text);
    /// <summary>Creates a comment in this mutable document.</summary>
    public HtmlNode CreateComment(string text) => CreateDataNode(HtmlNodeKind.Comment, text);
    /// <summary>Creates a detached fragment in this mutable document.</summary>
    public HtmlNode CreateFragment() => CreateDataNode(HtmlNodeKind.DocumentFragment, string.Empty);
    /// <summary>Creates a document type with the supplied identifiers.</summary>
    public HtmlDocumentType CreateDocumentType(string name, string publicIdentifier = "", string systemIdentifier = "") {
        EnsureMutable();
        return Register(new HtmlDocumentType(this, NextId(), name, publicIdentifier, systemIdentifier));
    }
    /// <summary>Creates or returns this template element's separate content fragment.</summary>
    public HtmlNode GetOrCreateTemplateContent(HtmlElement template) {
        EnsureMutable();
        if (template == null) throw new ArgumentNullException(nameof(template));
        if (!ReferenceEquals(template.Document, this) || template.NamespaceUri != HtmlElement.HtmlNamespace || template.LocalName != "template") throw new ArgumentException("An HTML template in this document is required.", nameof(template));
        if (template.TemplateContent == null) {
            template.TemplateContent = CreateFragment();
            template.TemplateContent.TemplateHost = template;
            Touch();
        }
        return template.TemplateContent;
    }
    /// <summary>Attaches an original UTF-16 source position while constructing a mutable tree.</summary>
    public void SetSourceIndex(HtmlNode node, int? sourceIndex) {
        EnsureMutable();
        if (node == null || !ReferenceEquals(node.Document, this)) throw new ArgumentException("A node in this document is required.", nameof(node));
        if (sourceIndex < 0) throw new ArgumentOutOfRangeException(nameof(sourceIndex));
        node.SourceIndex = sourceIndex;
        Touch();
    }
    /// <summary>Finds a node by its document-local identity, including detached nodes.</summary>
    public HtmlNode? GetNode(int nodeId) => _nodes.TryGetValue(nodeId, out HtmlNode? node) ? node : null;
    internal int RegisteredNodeCount => _nodes.Count;

    // Walk one sibling at a time so callers can stop on a budget without buffering a wide tree.
    internal IEnumerable<(HtmlNode Node, int Depth)> AttachedNodes() {
        yield return (this, 0);
        var pending = new Stack<(IEnumerator<HtmlNode> Nodes, int Depth)>();
        pending.Push((ChildNodes.GetEnumerator(), 0));
        try {
            while (pending.Count != 0) {
                var level = pending.Peek();
                if (!level.Nodes.MoveNext()) { level.Nodes.Dispose(); pending.Pop(); continue; }
                HtmlNode node = level.Nodes.Current;
                int depth = level.Depth + (node is HtmlElement ? 1 : 0);
                yield return (node, depth);
                if (node is HtmlElement element && element.TemplateContent != null)
                    pending.Push((((IEnumerable<HtmlNode>)new[] { element.TemplateContent }).GetEnumerator(), depth));
                pending.Push((node.ChildNodes.GetEnumerator(), depth));
            }
        } finally { while (pending.Count != 0) pending.Pop().Nodes.Dispose(); }
    }

    internal HtmlDocument CloneAttached(CancellationToken cancellationToken = default) {
        var clone = new HtmlDocument(Services, ProviderId, Mode);
        foreach (var entry in AttachedNodes()) {
            cancellationToken.ThrowIfCancellationRequested();
            HtmlNode source = entry.Node;
            if (ReferenceEquals(source, this)) continue;
            HtmlNode copy = CloneNode(source, clone, cancellationToken);
            if (source.TemplateHost != null) {
                var host = (HtmlElement)clone._nodes[source.TemplateHost.NodeId];
                host.TemplateContent = copy;
                copy.TemplateHost = host;
            } else clone._nodes[source.Parent!.NodeId].AppendChild(copy);
        }
        clone._nextNodeId = _nextNodeId;
        clone.Revision = Revision;
        return clone;
    }
    /// <summary>Freezes this instance. Every retained node handle becomes read-only.</summary>
    public HtmlDocument Freeze() { IsReadOnly = true; return this; }
    /// <summary>Creates an independent mutable clone, preserving node IDs and source positions.</summary>
    public HtmlDocument Clone() {
        var clone = new HtmlDocument(Services, ProviderId, Mode);
        foreach (HtmlNode node in _nodes.Values) if (!ReferenceEquals(node, this)) CloneNode(node, clone);
        foreach (HtmlNode source in _nodes.Values) {
            HtmlNode target = clone._nodes[source.NodeId];
            foreach (HtmlNode child in source.ChildNodes) {
                target.AppendChild(clone._nodes[child.NodeId]);
            }
            if (source is HtmlElement element && element.TemplateContent != null && target is HtmlElement targetElement) {
                HtmlNode content = clone._nodes[element.TemplateContent.NodeId];
                targetElement.TemplateContent = content;
                content.TemplateHost = targetElement;
            }
        }
        clone._nextNodeId = _nextNodeId;
        clone.Revision = Revision;
        return clone;
    }
    /// <summary>Applies edits to an independent clone and freezes the returned snapshot.</summary>
    public HtmlDocument Edit(Action<HtmlDocument> edit) {
        if (edit == null) throw new ArgumentNullException(nameof(edit));
        HtmlDocument clone = Clone();
        try { edit(clone); return clone.Freeze(); }
        finally { clone.Freeze(); }
    }
    internal void EnsureMutable() { if (IsReadOnly) throw new InvalidOperationException("This document is an immutable snapshot. Use Edit or Clone to change it."); }
    internal void Touch() { EnsureMutable(); Revision = checked(Revision + 1); }
    private int NextId() => checked(++_nextNodeId);
    private T Register<T>(T node) where T : HtmlNode { _nodes.Add(node.NodeId, node); Touch(); return node; }
    private HtmlNode CreateDataNode(HtmlNodeKind kind, string text) {
        EnsureMutable();
        return Register(new HtmlNode(this, NextId(), kind, text ?? throw new ArgumentNullException(nameof(text))));
    }
    private static HtmlNode CloneNode(HtmlNode source, HtmlDocument clone, CancellationToken cancellationToken = default) {
        HtmlNode copy;
        if (source is HtmlElement element) {
            var newElement = new HtmlElement(clone, source.NodeId, element.LocalName, element.NamespaceUri, element.Prefix);
            foreach (HtmlAttribute attribute in element.Attributes) {
                cancellationToken.ThrowIfCancellationRequested();
                newElement.SetAttribute(attribute.Name, attribute.Value, attribute.NamespaceUri);
            }
            copy = newElement;
        } else if (source is HtmlDocumentType type) copy = new HtmlDocumentType(clone, source.NodeId, type.Name, type.PublicIdentifier, type.SystemIdentifier);
        else copy = new HtmlNode(clone, source.NodeId, source.Kind, source.Data);
        copy.SourceIndex = source.SourceIndex;
        return clone.Register(copy);
    }
}

/// <summary>An owned document type declaration.</summary>
public sealed class HtmlDocumentType : HtmlNode {
    internal HtmlDocumentType(HtmlDocument document, int id, string name, string publicIdentifier, string systemIdentifier) : base(document, id, HtmlNodeKind.DocumentType) {
        Name = name ?? throw new ArgumentNullException(nameof(name));
        PublicIdentifier = publicIdentifier ?? throw new ArgumentNullException(nameof(publicIdentifier));
        SystemIdentifier = systemIdentifier ?? throw new ArgumentNullException(nameof(systemIdentifier));
    }
    /// <summary>Declared document type name.</summary>
    public string Name { get; }
    /// <summary>Public identifier, or an empty string.</summary>
    public string PublicIdentifier { get; }
    /// <summary>System identifier, or an empty string.</summary>
    public string SystemIdentifier { get; }
    /// <inheritdoc />
    public override string NodeName => Name;
}
