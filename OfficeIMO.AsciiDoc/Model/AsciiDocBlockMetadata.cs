namespace OfficeIMO.AsciiDoc;

internal interface IAsciiDocBlockMetadata {
    AsciiDocBlock? Target { get; set; }
}

/// <summary>Source-preserving block title line such as <c>.Example</c>.</summary>
public sealed class AsciiDocBlockTitle : AsciiDocBlock, IAsciiDocBlockMetadata {
    private string _title;
    private bool _titleWasAssigned;

    internal AsciiDocBlockTitle(AsciiDocSyntaxNode syntax, string title, AsciiDocInlineSequence inlines, string trailingLineEnding)
        : base(syntax, trailingLineEnding) {
        _title = title;
        Inlines = inlines;
    }

    /// <summary>Title without the leading dot.</summary>
    public string Title {
        get => !_titleWasAssigned && Inlines.IsModified ? Inlines.ToAsciiDoc() : _title;
        set {
            string normalized = value ?? string.Empty;
            AsciiDocText.EnsureSingleLine(normalized, nameof(value));
            if (SetValue(ref _title, normalized)) _titleWasAssigned = true;
        }
    }

    /// <summary>Typed inline title content.</summary>
    public AsciiDocInlineSequence Inlines { get; }

    /// <summary>Block this title describes, or null when detached.</summary>
    public AsciiDocBlock? Target { get; internal set; }

    AsciiDocBlock? IAsciiDocBlockMetadata.Target { get => Target; set => Target = value; }

    /// <inheritdoc />
    public override bool IsModified => base.IsModified || Inlines.IsModified;

    internal override string WriteCore(AsciiDocWriterContext context) =>
        "." + (_titleWasAssigned ? _title : Inlines.Write(context)) + EffectiveTrailingLineEnding(context);
}

/// <summary>Source-preserving block anchor line such as <c>[[id,Reference text]]</c>.</summary>
public sealed class AsciiDocBlockAnchor : AsciiDocBlock, IAsciiDocBlockMetadata {
    private string _id;
    private string? _referenceText;

    internal AsciiDocBlockAnchor(AsciiDocSyntaxNode syntax, string id, string? referenceText, string trailingLineEnding)
        : base(syntax, trailingLineEnding) {
        _id = id;
        _referenceText = referenceText;
    }

    /// <summary>Anchor ID.</summary>
    public string Id {
        get => _id;
        set { string normalized = value ?? string.Empty; AsciiDocText.EnsureSingleLine(normalized, nameof(value)); SetValue(ref _id, normalized); }
    }

    /// <summary>Optional reference text.</summary>
    public string? ReferenceText {
        get => _referenceText;
        set { if (value != null) AsciiDocText.EnsureSingleLine(value, nameof(value)); SetValue(ref _referenceText, value); }
    }

    /// <summary>Block this anchor identifies, or null when detached.</summary>
    public AsciiDocBlock? Target { get; internal set; }

    AsciiDocBlock? IAsciiDocBlockMetadata.Target { get => Target; set => Target = value; }

    internal override string WriteCore(AsciiDocWriterContext context) =>
        "[[" + Id + (ReferenceText == null ? string.Empty : "," + ReferenceText) + "]]" + EffectiveTrailingLineEnding(context);
}
