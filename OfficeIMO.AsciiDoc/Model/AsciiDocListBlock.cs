namespace OfficeIMO.AsciiDoc;

/// <summary>AsciiDoc list kind recognized by the parser.</summary>
public enum AsciiDocListKind {
    /// <summary>Unordered list.</summary>
    Unordered = 0,
    /// <summary>Ordered list.</summary>
    Ordered = 1,
    /// <summary>Numbered explanations corresponding to callout markers in a preceding verbatim block.</summary>
    Callout = 2
}

/// <summary>Source-backed item within an AsciiDoc list.</summary>
public sealed class AsciiDocListItem {
    private readonly AsciiDocEditableInlineContent _content;
    private readonly List<AsciiDocBlock> _attachedBlocks = new List<AsciiDocBlock>();

    internal AsciiDocListItem(
        AsciiDocSyntaxNode syntax,
        AsciiDocListKind kind,
        string marker,
        int depth,
        string text,
        AsciiDocInlineSequence inlines,
        string trailingLineEnding) {
        Syntax = syntax;
        Kind = kind;
        Marker = marker;
        Depth = depth;
        _content = new AsciiDocEditableInlineContent(text, inlines);
        TrailingLineEnding = trailingLineEnding;
    }

    /// <summary>Lossless syntax for the complete item line.</summary>
    public AsciiDocSyntaxNode Syntax { get; }

    /// <summary>List kind.</summary>
    public AsciiDocListKind Kind { get; }

    /// <summary>Original list marker.</summary>
    public string Marker { get; }

    /// <summary>Marker-derived nesting depth.</summary>
    public int Depth { get; }

    /// <summary>Item text after the marker.</summary>
    public string Text {
        get => _content.Text;
        set {
            string normalized = value ?? string.Empty;
            AsciiDocText.EnsureSingleLine(normalized, nameof(value));
            _content.Assign(normalized);
        }
    }

    /// <summary>Typed lossless inline content in the item.</summary>
    public AsciiDocInlineSequence Inlines => _content.Inlines;

    /// <summary>Compound blocks attached through list continuation markers.</summary>
    public IReadOnlyList<AsciiDocBlock> AttachedBlocks => _attachedBlocks;

    /// <summary>True when the item text was edited.</summary>
    public bool IsModified => _content.IsModified;

    internal string TrailingLineEnding { get; }

    internal void AddAttachedBlock(AsciiDocBlock block) => _attachedBlocks.Add(block);
    internal void RemoveAttachedBlock(AsciiDocBlock block) => _attachedBlocks.Remove(block);

    internal string Write(AsciiDocWriterContext context) {
        if (context.Mode == AsciiDocWriterMode.Preserve && !IsModified) return Syntax.OriginalText;
        string marker = context.Mode == AsciiDocWriterMode.Preserve
            ? Marker
            : (Kind == AsciiDocListKind.Callout ? Marker : Kind == AsciiDocListKind.Ordered ? new string('.', Depth) : new string('*', Depth));
        string ending = context.Mode == AsciiDocWriterMode.Preserve
            ? TrailingLineEnding
            : (TrailingLineEnding.Length == 0 ? string.Empty : context.LineEnding);
        return marker + " " + _content.Write(context) + ending;
    }
}

/// <summary>Contiguous ordered or unordered AsciiDoc list.</summary>
public sealed class AsciiDocListBlock : AsciiDocBlock {
    private readonly IReadOnlyList<AsciiDocListItem> _items;

    internal AsciiDocListBlock(
        AsciiDocSyntaxNode syntax,
        AsciiDocListKind kind,
        IReadOnlyList<AsciiDocListItem> items,
        string trailingLineEnding)
        : base(syntax, trailingLineEnding) {
        Kind = kind;
        _items = items;
    }

    /// <summary>Ordered or unordered list kind.</summary>
    public AsciiDocListKind Kind { get; }

    /// <summary>Source-backed items in document order.</summary>
    public IReadOnlyList<AsciiDocListItem> Items => _items;

    /// <inheritdoc />
    public override bool IsModified => base.IsModified || Items.Any(static item => item.IsModified);

    internal override string WriteCore(AsciiDocWriterContext context) {
        var builder = new StringBuilder();
        for (int index = 0; index < Items.Count; index++) builder.Append(Items[index].Write(context));
        return builder.ToString();
    }
}
