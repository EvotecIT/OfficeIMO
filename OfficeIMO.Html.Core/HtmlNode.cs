using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;

namespace OfficeIMO.Html.Dom;

/// <summary>An owned node. Parsed snapshots are read-only; edits operate on an independent mutable clone.</summary>
public class HtmlNode {
    private readonly HtmlDocument? _document;
    private readonly List<HtmlNode> _children = new List<HtmlNode>();
    private readonly ReadOnlyCollection<HtmlNode> _childrenView;
    private string _data;

    internal HtmlNode(HtmlDocument? document, int nodeId, HtmlNodeKind kind, string data = "") {
        _document = document;
        NodeId = nodeId;
        Kind = kind;
        _data = data;
        _childrenView = _children.AsReadOnly();
    }

    /// <summary>Document owning this node, including detached nodes.</summary>
    public HtmlDocument Document => _document ?? (HtmlDocument)this;
    /// <summary>Stable node identity within this document's edit lineage.</summary>
    public int NodeId { get; }
    /// <summary>Structural node category.</summary>
    public HtmlNodeKind Kind { get; }
    /// <summary>Parent node, or null for the document and detached nodes.</summary>
    public HtmlNode? Parent { get; private set; }
    internal HtmlElement? TemplateHost { get; set; }
    /// <summary>Parent element when the immediate parent is an element.</summary>
    public HtmlElement? ParentElement => Parent as HtmlElement;
    /// <summary>Ordered child nodes. The collection cannot be mutated directly.</summary>
    public IReadOnlyList<HtmlNode> ChildNodes => _childrenView;
    /// <summary>Direct child elements in source/tree order.</summary>
    public IEnumerable<HtmlElement> Children => _children.OfType<HtmlElement>();
    /// <summary>First child, if any.</summary>
    public HtmlNode? FirstChild => _children.Count == 0 ? null : _children[0];
    /// <summary>Next sibling, if any.</summary>
    public HtmlNode? NextSibling => GetSibling(1);
    /// <summary>Previous sibling, if any.</summary>
    public HtmlNode? PreviousSibling => GetSibling(-1);
    /// <summary>Original UTF-16 source offset, or null for an implied or authored node.</summary>
    public int? SourceIndex { get; internal set; }
    /// <summary>Node name for non-element nodes; elements expose their local name.</summary>
    public virtual string NodeName => Kind == HtmlNodeKind.Text ? "#text" : Kind == HtmlNodeKind.Comment ? "#comment" : Kind.ToString();
    /// <summary>Text or comment data. Container nodes aggregate descendant text.</summary>
    public string TextContent {
        get {
            if (Kind == HtmlNodeKind.Text || Kind == HtmlNodeKind.Comment) return _data;
            var text = new StringBuilder();
            foreach (HtmlNode node in Descendants()) if (node.Kind == HtmlNodeKind.Text) text.Append(node._data);
            return text.ToString();
        }
        set {
            Document.EnsureMutable();
            if (value == null) throw new ArgumentNullException(nameof(value));
            if (Kind == HtmlNodeKind.DocumentType || Kind == HtmlNodeKind.Document) throw new InvalidOperationException("Set text on a text, comment, element or fragment node.");
            if (Kind == HtmlNodeKind.Text || Kind == HtmlNodeKind.Comment) _data = value;
            else {
                foreach (HtmlNode child in _children) child.Parent = null;
                _children.Clear();
                if (value.Length != 0) AppendChild(Document.CreateTextNode(value));
            }
            Document.Touch();
        }
    }
    /// <summary>HTML serialization of this node.</summary>
    public string OuterHtml => Document.Services.Serialize(this);
    /// <summary>HTML serialization of this node's children.</summary>
    public string InnerHtml => Document.Services.Serialize(this, true);

    /// <summary>Appends or moves a node in the same mutable document. Cycles and cross-document moves are rejected.</summary>
    public HtmlNode AppendChild(HtmlNode child) {
        if (child == null) throw new ArgumentNullException(nameof(child));
        Document.EnsureMutable();
        if (Kind != HtmlNodeKind.Document && Kind != HtmlNodeKind.Element && Kind != HtmlNodeKind.DocumentFragment) throw new InvalidOperationException("This node cannot contain children.");
        if (!ReferenceEquals(Document, child.Document)) throw new ArgumentException("The child belongs to another document.", nameof(child));
        for (HtmlNode? ancestor = this; ancestor != null; ancestor = ancestor.Parent ?? ancestor.TemplateHost) if (ReferenceEquals(ancestor, child)) throw new ArgumentException("The operation would create a tree cycle.", nameof(child));
        if (child.TemplateHost != null) throw new ArgumentException("Template content cannot be moved as a normal child node.", nameof(child));
        if (child.Kind == HtmlNodeKind.Document) throw new ArgumentException("A document cannot be appended.", nameof(child));
        HtmlNode[] additions = child.Kind == HtmlNodeKind.DocumentFragment ? child.ChildNodes.ToArray() : new[] { child };
        if (Kind == HtmlNodeKind.Document) {
            HtmlNode[] result = _children.Except(additions).Concat(additions).ToArray();
            if (result.Any(node => node.Kind == HtmlNodeKind.Text) || result.Count(node => node.Kind == HtmlNodeKind.Element) > 1 || result.Count(node => node.Kind == HtmlNodeKind.DocumentType) > 1)
                throw new ArgumentException("A document accepts one root element, one document type and comments, without direct text.", nameof(child));
            int typeIndex = Array.FindIndex(result, node => node.Kind == HtmlNodeKind.DocumentType);
            int elementIndex = Array.FindIndex(result, node => node.Kind == HtmlNodeKind.Element);
            if (typeIndex >= 0 && elementIndex >= 0 && typeIndex > elementIndex) throw new ArgumentException("The document type must precede the root element.", nameof(child));
        } else if (additions.Any(node => node.Kind == HtmlNodeKind.DocumentType)) throw new ArgumentException("A document type must be a direct document child.", nameof(child));
        if (child.Kind == HtmlNodeKind.DocumentFragment) {
            foreach (HtmlNode addition in additions) AppendChild(addition);
            return child;
        }
        child.Parent?._children.Remove(child);
        child.Parent = this;
        _children.Add(child);
        Document.Touch();
        return child;
    }

    /// <summary>Detaches this node without affecting sibling nodes.</summary>
    public void Remove() {
        Document.EnsureMutable();
        if (Parent == null) return;
        Parent._children.Remove(this);
        Parent = null;
        Document.Touch();
    }

    /// <summary>Enumerates descendants in tree order without recursive traversal.</summary>
    public IEnumerable<HtmlNode> Descendants() {
        var pending = new Stack<HtmlNode>();
        for (int index = _children.Count - 1; index >= 0; index--) pending.Push(_children[index]);
        while (pending.Count != 0) {
            HtmlNode current = pending.Pop();
            yield return current;
            for (int index = current._children.Count - 1; index >= 0; index--) pending.Push(current._children[index]);
        }
    }

    /// <summary>Finds the first descendant matching a CSS selector.</summary>
    public HtmlElement? QuerySelector(string selector) => QuerySelectorAll(selector).FirstOrDefault();
    /// <summary>Finds matching descendants using this document's selector provider.</summary>
    public IReadOnlyList<HtmlElement> QuerySelectorAll(string selector) {
        if (selector == null) throw new ArgumentNullException(nameof(selector));
        return Document.Services.QuerySelectorAll(this, selector);
    }

    private HtmlNode? GetSibling(int offset) {
        if (Parent == null) return null;
        int index = Parent._children.IndexOf(this) + offset;
        return index >= 0 && index < Parent._children.Count ? Parent._children[index] : null;
    }

    internal string Data => _data;
}
