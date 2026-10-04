namespace OfficeIMO.AsciiDoc;

/// <summary>One source-backed description list item.</summary>
public sealed class AsciiDocDescriptionListItem {
    private readonly AsciiDocEditableInlineContent _term;
    private readonly AsciiDocEditableInlineContent _description;

    internal AsciiDocDescriptionListItem(
        AsciiDocSyntaxNode syntax,
        string marker,
        string term,
        string description,
        AsciiDocInlineSequence termInlines,
        AsciiDocInlineSequence descriptionInlines,
        string trailingLineEnding) {
        Syntax = syntax;
        Marker = marker;
        _term = new AsciiDocEditableInlineContent(term, termInlines);
        _description = new AsciiDocEditableInlineContent(description, descriptionInlines);
        TrailingLineEnding = trailingLineEnding;
    }

    /// <summary>Lossless item syntax.</summary>
    public AsciiDocSyntaxNode Syntax { get; }

    /// <summary>Original repeated-colon marker.</summary>
    public string Marker { get; }

    /// <summary>Marker-derived nesting depth.</summary>
    public int Depth => Math.Max(1, Marker.Length - 1);

    /// <summary>Term text.</summary>
    public string Term {
        get => _term.Text;
        set { string normalized = value ?? string.Empty; AsciiDocText.EnsureSingleLine(normalized, nameof(value)); _term.Assign(normalized); }
    }

    /// <summary>Definition text on the item line.</summary>
    public string Description {
        get => _description.Text;
        set { string normalized = value ?? string.Empty; AsciiDocText.EnsureSingleLine(normalized, nameof(value)); _description.Assign(normalized); }
    }

    /// <summary>Typed term inlines.</summary>
    public AsciiDocInlineSequence TermInlines => _term.Inlines;

    /// <summary>Typed definition inlines.</summary>
    public AsciiDocInlineSequence DescriptionInlines => _description.Inlines;

    /// <summary>True when text or nested inline content changed.</summary>
    public bool IsModified => _term.IsModified || _description.IsModified;

    internal string TrailingLineEnding { get; }

    internal string Write(AsciiDocWriterContext context) {
        if (context.Mode == AsciiDocWriterMode.Preserve && !IsModified) return Syntax.OriginalText;
        string term = _term.Write(context);
        string description = _description.Write(context);
        string ending = context.Mode == AsciiDocWriterMode.Preserve ? TrailingLineEnding : (TrailingLineEnding.Length == 0 ? string.Empty : context.LineEnding);
        return term + Marker + (description.Length == 0 ? string.Empty : " " + description) + ending;
    }
}

/// <summary>Contiguous AsciiDoc description list.</summary>
public sealed class AsciiDocDescriptionListBlock : AsciiDocBlock {
    private readonly IReadOnlyList<AsciiDocDescriptionListItem> _items;

    internal AsciiDocDescriptionListBlock(AsciiDocSyntaxNode syntax, IReadOnlyList<AsciiDocDescriptionListItem> items, string trailingLineEnding)
        : base(syntax, trailingLineEnding) {
        _items = items;
    }

    /// <summary>Items in source order.</summary>
    public IReadOnlyList<AsciiDocDescriptionListItem> Items => _items;

    /// <inheritdoc />
    public override bool IsModified => base.IsModified || Items.Any(static item => item.IsModified);

    internal override string WriteCore(AsciiDocWriterContext context) {
        var output = new StringBuilder();
        for (int index = 0; index < Items.Count; index++) output.Append(Items[index].Write(context));
        return output.ToString();
    }
}
