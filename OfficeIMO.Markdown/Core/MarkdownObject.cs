namespace OfficeIMO.Markdown;

/// <summary>
/// Base type for the navigable OfficeIMO.Markdown object tree.
/// </summary>
public abstract class MarkdownObject {
    // Parsed objects can use their already-retained syntax node as immutable source metadata.
    // Objects edited or constructed independently fall back to a writable metadata holder.
    private object? _metadata;

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
    public MarkdownSourceSpan? SourceSpan {
        get => _metadata switch {
            MarkdownSyntaxNode syntaxNode => syntaxNode.SourceSpan,
            MarkdownObjectMetadata metadata => metadata.SourceSpan,
            MarkdownSourceSpan sourceSpan => sourceSpan,
            _ => null
        };
        internal set {
            if (value.HasValue) {
                if (_metadata == null || _metadata is MarkdownSourceSpan) {
                    _metadata = value.Value;
                } else {
                    WritableMetadata.SourceSpan = value;
                }
            } else if (_metadata is MarkdownObjectMetadata metadata) {
                metadata.SourceSpan = null;
            } else if (_metadata is MarkdownSourceSpan) {
                _metadata = null;
            } else if (_metadata is MarkdownSyntaxNode syntaxNode) {
                _metadata = syntaxNode.Attributes.IsEmpty
                    ? null
                    : new MarkdownObjectMetadata { Attributes = syntaxNode.Attributes };
            }
        }
    }

    /// <summary>Generic Markdown attributes associated with this node.</summary>
    public MarkdownAttributeSet Attributes => _metadata switch {
        MarkdownSyntaxNode syntaxNode => syntaxNode.Attributes,
        MarkdownObjectMetadata metadata => metadata.Attributes,
        _ => MarkdownAttributeSet.Empty
    };

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

    internal void SetAttributes(MarkdownAttributeSet? attributes) {
        if (attributes == null || attributes.IsEmpty) {
            if (_metadata is MarkdownObjectMetadata metadata) {
                metadata.Attributes = MarkdownAttributeSet.Empty;
            } else if (_metadata is MarkdownSyntaxNode syntaxNode) {
                _metadata = syntaxNode.SourceSpan.HasValue
                    ? new MarkdownObjectMetadata { SourceSpan = syntaxNode.SourceSpan }
                    : null;
            }
            return;
        }

        WritableMetadata.Attributes = attributes;
    }

    internal void BindSyntaxNode(MarkdownSyntaxNode syntaxNode) {
        if (syntaxNode == null) {
            throw new ArgumentNullException(nameof(syntaxNode));
        }

        MarkdownAttributeSet existingAttributes = Attributes;
        _metadata = syntaxNode.Attributes.IsEmpty && !existingAttributes.IsEmpty
            ? new MarkdownObjectMetadata {
                SourceSpan = syntaxNode.SourceSpan,
                Attributes = existingAttributes
            }
            : syntaxNode;
    }

    internal MarkdownSyntaxNode? BoundSyntaxNode => _metadata as MarkdownSyntaxNode;

    private MarkdownObjectMetadata WritableMetadata {
        get {
            if (_metadata is MarkdownObjectMetadata metadata) {
                return metadata;
            }

            var created = _metadata switch {
                MarkdownSyntaxNode syntaxNode => new MarkdownObjectMetadata {
                    SourceSpan = syntaxNode.SourceSpan,
                    Attributes = syntaxNode.Attributes
                },
                MarkdownSourceSpan sourceSpan => new MarkdownObjectMetadata {
                    SourceSpan = sourceSpan
                },
                _ => new MarkdownObjectMetadata()
            };
            _metadata = created;
            return created;
        }
    }
}

internal sealed class MarkdownObjectMetadata {
    internal MarkdownSourceSpan? SourceSpan;
    internal MarkdownAttributeSet Attributes = MarkdownAttributeSet.Empty;
}

/// <summary>
/// Base type for markdown blocks that participate in the object tree.
/// </summary>
public abstract class MarkdownBlock : MarkdownObject { }

/// <summary>
/// Base type for markdown inlines that participate in the object tree.
/// </summary>
public abstract class MarkdownInline : MarkdownObject, IMarkdownInline { }

internal interface IMarkdownInlineAuxiliarySyntaxMetadataOwner {
    MarkdownInlineAuxiliarySyntaxMetadata? AuxiliarySyntaxMetadata { get; set; }
}

internal sealed class MarkdownInlineAuxiliarySyntaxMetadata {
    private MarkdownSourceSpan _openingMarkerSpan;
    private MarkdownSourceSpan _closingMarkerSpan;
    private MarkdownInlineAuxiliarySyntaxRareMetadata? _rare;

    internal string OpeningMarker = string.Empty;
    internal string ClosingMarker = string.Empty;

    internal string? AutolinkLiteral {
        get => _rare?.AutolinkLiteral;
        set {
            if (string.IsNullOrEmpty(value) && _rare == null) {
                return;
            }

            (_rare ??= new MarkdownInlineAuxiliarySyntaxRareMetadata()).AutolinkLiteral = value;
        }
    }

    internal MarkdownSourceSpan? OpeningMarkerSpan {
        get => _openingMarkerSpan.StartLine == 0 ? null : _openingMarkerSpan;
        set => _openingMarkerSpan = value.GetValueOrDefault();
    }

    internal string SeparatorMarker {
        get => _rare?.SeparatorMarker ?? string.Empty;
        set {
            if (string.IsNullOrEmpty(value) && _rare == null) {
                return;
            }

            (_rare ??= new MarkdownInlineAuxiliarySyntaxRareMetadata()).SeparatorMarker = value ?? string.Empty;
        }
    }

    internal MarkdownSourceSpan? SeparatorMarkerSpan {
        get {
            MarkdownSourceSpan span = _rare?.SeparatorMarkerSpan ?? default;
            return span.StartLine == 0 ? null : span;
        }
        set {
            if (!value.HasValue && _rare == null) {
                return;
            }

            (_rare ??= new MarkdownInlineAuxiliarySyntaxRareMetadata()).SeparatorMarkerSpan = value.GetValueOrDefault();
        }
    }

    internal MarkdownSourceSpan? ClosingMarkerSpan {
        get => _closingMarkerSpan.StartLine == 0 ? null : _closingMarkerSpan;
        set => _closingMarkerSpan = value.GetValueOrDefault();
    }
}

internal sealed class MarkdownInlineAuxiliarySyntaxRareMetadata {
    internal string? AutolinkLiteral;
    internal string SeparatorMarker = string.Empty;
    internal MarkdownSourceSpan SeparatorMarkerSpan;
}
