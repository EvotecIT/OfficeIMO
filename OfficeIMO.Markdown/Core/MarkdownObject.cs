namespace OfficeIMO.Markdown;

/// <summary>
/// Base type for the navigable OfficeIMO.Markdown object tree.
/// </summary>
public abstract class MarkdownObject {
    // Parsed objects can use their already-retained syntax node as immutable source metadata.
    // Objects edited or constructed independently fall back to a writable metadata holder.
    // Objects of a parsed document whose source information is not captured yet hold the
    // document's deferred capture, which supplies their syntax node on first use.
    private object? _metadata;

    /// <summary>
    /// Captures the parsed source information of this object's document when it was deferred.
    /// Every member that reads or writes source information calls this first.
    /// </summary>
    internal void EnsureSourceInformation() {
        if (System.Threading.Volatile.Read(ref _metadata) is MarkdownDeferredSourceCapture deferred) {
            deferred.Capture();
        }
    }

    /// <summary>Reads a source-information field of this object, capturing deferred source information first.</summary>
    private protected T ReadSourceInformation<T>(ref T storage) {
        EnsureSourceInformation();
        return storage;
    }

    /// <summary>Writes a source-information field of this object, capturing deferred source information first.</summary>
    private protected void WriteSourceInformation<T>(ref T storage, T value) {
        EnsureSourceInformation();
        storage = value;
    }

    /// <summary>
    /// Reads one of several optional source spans kept in a side array. Types with several such
    /// spans keep them out of line so objects from a parse without source capture stay small.
    /// </summary>
    private protected MarkdownSourceSpan? ReadSourceSpanSlot(ref MarkdownSourceSpan?[]? slots, int index) {
        EnsureSourceInformation();
        MarkdownSourceSpan?[]? current = slots;
        return current == null ? null : current[index];
    }

    /// <summary>Writes one of several optional source spans, creating the side array on the first value.</summary>
    private protected void WriteSourceSpanSlot(ref MarkdownSourceSpan?[]? slots, int index, int slotCount, MarkdownSourceSpan? value) {
        EnsureSourceInformation();
        if (slots == null) {
            if (!value.HasValue) {
                return;
            }

            slots = new MarkdownSourceSpan?[slotCount];
        }

        slots[index] = value;
    }

    internal bool HasMetadata => _metadata != null;

    internal void DeferSourceInformation(MarkdownDeferredSourceCapture capture) => _metadata = capture;

    // While source capture attaches a source-backed parse to a document, each object of that parse
    // points at the document's object it describes, and gets its own metadata back at the end.
    internal object? ForwardCapturedObject(MarkdownObject target) {
        object? metadata = _metadata;
        _metadata = target;
        return metadata;
    }

    internal void RestoreCapturedMetadata(object? metadata) => _metadata = metadata;

    // Completes a deferred object once everything about it is attached; a reader that sees the new
    // metadata also sees the rest of its source information.
    internal void PublishCapturedMetadata(object? metadata) => System.Threading.Volatile.Write(ref _metadata, metadata);

    internal MarkdownObject? CapturedObjectTarget => _metadata as MarkdownObject;

    internal void ClearDeferredSourceInformation() {
        if (_metadata is MarkdownDeferredSourceCapture) {
            _metadata = null;
        }
    }

    /// <summary>
    /// Copies the source information that a source-backed parse records on an object of the same type,
    /// beyond its syntax node. Types holding such information override this.
    /// </summary>
    internal virtual void AdoptCapturedSourceInformation(MarkdownObject captured, MarkdownCapturedObjectMap map) {
    }

    /// <summary>Parent node in the markdown object tree, or <c>null</c> for the document root.</summary>
    public MarkdownObject? Parent { get; private set; }

    /// <summary>Containing document for this node when attached to a tree.</summary>
    public MarkdownDocument? Document => this as MarkdownDocument ?? Parent?.Document;

    /// <summary>Root node for this object's tree.</summary>
    public MarkdownObject Root => Parent?.Root ?? this;

    // Stored as -1 when there is no index, which keeps the field four bytes wide: derived types
    // place their own small fields in the space a nullable index would take.
    private int _indexInParent = -1;

    /// <summary>Zero-based index within the parent when this node is attached to a tree.</summary>
    public int? IndexInParent => _indexInParent < 0 ? null : _indexInParent;

    /// <summary>Previous sibling in the parent container when available.</summary>
    public MarkdownObject? PreviousSibling { get; private set; }

    /// <summary>Next sibling in the parent container when available.</summary>
    public MarkdownObject? NextSibling { get; private set; }

    /// <summary>Source span mapped from the syntax tree when available.</summary>
    public MarkdownSourceSpan? SourceSpan {
        get {
            EnsureSourceInformation();
            return _metadata switch {
                MarkdownSyntaxNode syntaxNode => syntaxNode.SourceSpan,
                MarkdownObjectMetadata metadata => metadata.SourceSpan,
                MarkdownSourceSpan sourceSpan => sourceSpan,
                _ => null
            };
        }
        internal set {
            EnsureSourceInformation();
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

    // An object whose source capture is deferred has no attributes: a parse that finds generic
    // attributes captures source information while parsing.
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
        ArgumentNullException.ThrowIfNull(visitor);

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
        _indexInParent = indexInParent ?? -1;
        PreviousSibling = previousSibling;
        NextSibling = nextSibling;
    }

    internal void SetAttributes(MarkdownAttributeSet? attributes) {
        EnsureSourceInformation();
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
        ArgumentNullException.ThrowIfNull(syntaxNode);

        EnsureSourceInformation();
        MarkdownAttributeSet existingAttributes = Attributes;
        _metadata = syntaxNode.Attributes.IsEmpty && !existingAttributes.IsEmpty
            ? new MarkdownObjectMetadata {
                SourceSpan = syntaxNode.SourceSpan,
                Attributes = existingAttributes
            }
            : syntaxNode;
    }

    // Rendering reads the parse state it needs without capturing deferred source information: the
    // semantic parse leaves what writing reads in the same rendered form as a source-backed parse.
    internal MarkdownSyntaxNode? BoundSyntaxNodeWithoutCapture => _metadata as MarkdownSyntaxNode;

    internal MarkdownSyntaxNode? BoundSyntaxNode {
        get {
            EnsureSourceInformation();
            return _metadata as MarkdownSyntaxNode;
        }
    }

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
