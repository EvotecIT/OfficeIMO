namespace OfficeIMO.AsciiDoc;

/// <summary>One source-backed description list item.</summary>
public sealed class AsciiDocDescriptionListItem {
    private string _term;
    private string _description;
    private bool _termAssigned;
    private bool _descriptionAssigned;

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
        _term = term;
        _description = description;
        TermInlines = termInlines;
        DescriptionInlines = descriptionInlines;
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
        get => !_termAssigned && TermInlines.IsModified ? TermInlines.ToAsciiDoc() : _term;
        set { string normalized = value ?? string.Empty; AsciiDocText.EnsureSingleLine(normalized, nameof(value)); if (_term != normalized) { _term = normalized; _termAssigned = true; } }
    }

    /// <summary>Definition text on the item line.</summary>
    public string Description {
        get => !_descriptionAssigned && DescriptionInlines.IsModified ? DescriptionInlines.ToAsciiDoc() : _description;
        set { string normalized = value ?? string.Empty; AsciiDocText.EnsureSingleLine(normalized, nameof(value)); if (_description != normalized) { _description = normalized; _descriptionAssigned = true; } }
    }

    /// <summary>Typed term inlines.</summary>
    public AsciiDocInlineSequence TermInlines { get; }

    /// <summary>Typed definition inlines.</summary>
    public AsciiDocInlineSequence DescriptionInlines { get; }

    /// <summary>True when text or nested inline content changed.</summary>
    public bool IsModified => _termAssigned || _descriptionAssigned || TermInlines.IsModified || DescriptionInlines.IsModified;

    internal string TrailingLineEnding { get; }

    internal string Write(AsciiDocWriterContext context) {
        if (context.Mode == AsciiDocWriterMode.Preserve && !IsModified) return Syntax.OriginalText;
        string term = _termAssigned ? _term : TermInlines.Write(context);
        string description = _descriptionAssigned ? _description : DescriptionInlines.Write(context);
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
