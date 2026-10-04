namespace OfficeIMO.AsciiDoc;

/// <summary>AsciiDoc document title or section heading.</summary>
public sealed class AsciiDocHeading : AsciiDocBlock {
    private readonly AsciiDocEditableInlineContent _content;

    internal AsciiDocHeading(
        AsciiDocSyntaxNode syntax,
        string marker,
        string title,
        bool isDocumentTitle,
        AsciiDocInlineSequence inlines,
        string trailingLineEnding)
        : base(syntax, trailingLineEnding) {
        Marker = marker;
        _content = new AsciiDocEditableInlineContent(title, inlines);
        IsDocumentTitle = isDocumentTitle;
    }

    /// <summary>Original equals-sign marker.</summary>
    public string Marker { get; }

    /// <summary>Number of equals signs in the heading marker.</summary>
    public int MarkerLevel => Marker.Length;

    /// <summary>Logical section level. A document title has level 0.</summary>
    public int SectionLevel => IsDocumentTitle ? 0 : Math.Max(1, MarkerLevel - 1);

    /// <summary>True when this heading is the document title.</summary>
    public bool IsDocumentTitle { get; }

    /// <summary>Heading text.</summary>
    public string Title {
        get => _content.Text;
        set {
            string normalized = value ?? string.Empty;
            AsciiDocText.EnsureSingleLine(normalized, nameof(value));
            _content.Assign(normalized);
        }
    }

    /// <summary>Typed lossless inline content in the title.</summary>
    public AsciiDocInlineSequence Inlines => _content.Inlines;

    /// <inheritdoc />
    public override bool IsModified => base.IsModified || _content.IsModified;

    internal override string WriteCore(AsciiDocWriterContext context) =>
        Marker + " " + _content.Write(context) + EffectiveTrailingLineEnding(context);
}
