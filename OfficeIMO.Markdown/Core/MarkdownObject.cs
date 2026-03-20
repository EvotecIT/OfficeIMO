namespace OfficeIMO.Markdown;

/// <summary>
/// Base type for the navigable OfficeIMO.Markdown object tree.
/// </summary>
public abstract class MarkdownObject {
    /// <summary>Parent node in the markdown object tree, or <c>null</c> for the document root.</summary>
    public MarkdownObject? Parent { get; private set; }

    /// <summary>Containing document for this node when attached to a tree.</summary>
    public MarkdownDoc? Document => this as MarkdownDoc ?? Parent?.Document;

    /// <summary>Root node for this object's tree.</summary>
    public MarkdownObject Root => Parent?.Root ?? this;

    /// <summary>Zero-based index within the parent when this node is attached to a tree.</summary>
    public int? IndexInParent { get; private set; }

    /// <summary>Previous sibling in the parent container when available.</summary>
    public MarkdownObject? PreviousSibling { get; private set; }

    /// <summary>Next sibling in the parent container when available.</summary>
    public MarkdownObject? NextSibling { get; private set; }

    /// <summary>Source span mapped from the syntax tree when available.</summary>
    public MarkdownSourceSpan? SourceSpan { get; internal set; }

    /// <summary>Immediate child objects in document order.</summary>
    public IReadOnlyList<MarkdownObject> ChildObjects => MarkdownObjectTreeBinder.GetChildObjects(this);

    /// <summary>Dispatches this node to a visitor.</summary>
    public void Accept(MarkdownVisitor visitor) {
        if (visitor == null) {
            throw new ArgumentNullException(nameof(visitor));
        }

        visitor.Visit(this);
    }

    /// <summary>Enumerates ancestor nodes starting from the parent.</summary>
    public IEnumerable<MarkdownObject> Ancestors() {
        for (var current = Parent; current != null; current = current.Parent) {
            yield return current;
        }
    }

    /// <summary>Enumerates this node followed by its ancestors.</summary>
    public IEnumerable<MarkdownObject> AncestorsAndSelf() {
        for (MarkdownObject? current = this; current != null; current = current.Parent) {
            yield return current;
        }
    }

    /// <summary>Enumerates descendant nodes in depth-first order.</summary>
    public IEnumerable<MarkdownObject> Descendants() {
        var children = ChildObjects;
        for (int i = 0; i < children.Count; i++) {
            yield return children[i];

            foreach (var descendant in children[i].Descendants()) {
                yield return descendant;
            }
        }
    }

    /// <summary>Enumerates descendants of the requested node type.</summary>
    public IEnumerable<TObject> DescendantObjectsOfType<TObject>() where TObject : MarkdownObject {
        foreach (var descendant in Descendants()) {
            if (descendant is TObject typed) {
                yield return typed;
            }
        }
    }

    internal void SetTreePosition(
        MarkdownObject? parent,
        int? indexInParent,
        MarkdownObject? previousSibling,
        MarkdownObject? nextSibling) {
        Parent = parent;
        IndexInParent = indexInParent;
        PreviousSibling = previousSibling;
        NextSibling = nextSibling;
    }
}

/// <summary>
/// Base type for markdown blocks that participate in the object tree.
/// </summary>
public abstract class MarkdownBlock : MarkdownObject { }

/// <summary>
/// Base type for markdown inlines that participate in the object tree.
/// </summary>
public abstract class MarkdownInline : MarkdownObject, IMarkdownInline { }
