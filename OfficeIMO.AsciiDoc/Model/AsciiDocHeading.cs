namespace OfficeIMO.AsciiDoc;

/// <summary>AsciiDoc document title or section heading.</summary>
public sealed class AsciiDocHeading : AsciiDocBlock {
    private string _title;
    private bool _titleWasAssigned;

    internal AsciiDocHeading(
        AsciiDocSyntaxNode syntax,
        string marker,
        string title,
        bool isDocumentTitle,
        AsciiDocInlineSequence inlines,
        string trailingLineEnding)
        : base(syntax, trailingLineEnding) {
        Marker = marker;
        _title = title;
        IsDocumentTitle = isDocumentTitle;
        Inlines = inlines;
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
        get => !_titleWasAssigned && Inlines.IsModified ? Inlines.ToAsciiDoc() : _title;
        set {
            string normalized = value ?? string.Empty;
            AsciiDocText.EnsureSingleLine(normalized, nameof(value));
            if (SetValue(ref _title, normalized)) _titleWasAssigned = true;
        }
    }

    /// <summary>Typed lossless inline content in the title.</summary>
    public AsciiDocInlineSequence Inlines { get; }

    /// <inheritdoc />
    public override bool IsModified => base.IsModified || Inlines.IsModified;

    internal override string WriteCore(AsciiDocWriterContext context) =>
        Marker + " " + (_titleWasAssigned ? _title : Inlines.Write(context)) + EffectiveTrailingLineEnding(context);
}
