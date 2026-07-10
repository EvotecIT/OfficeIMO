namespace OfficeIMO.AsciiDoc;

/// <summary>Inline document attribute reference such as <c>{name}</c>.</summary>
public sealed class AsciiDocAttributeReferenceInline : AsciiDocInline {
    private string _name;

    internal AsciiDocAttributeReferenceInline(AsciiDocSyntaxNode syntax, string name) : base(syntax) {
        _name = name;
    }

    /// <summary>Referenced attribute name.</summary>
    public string Name {
        get => _name;
        set {
            string normalized = value ?? string.Empty;
            if (!AsciiDocText.IsAttributeName(normalized)) throw new ArgumentException("Invalid attribute name.", nameof(value));
            SetValue(ref _name, normalized);
        }
    }

    internal override string WriteCore(AsciiDocWriterContext context) => "{" + Name + "}";
}

/// <summary>Inline cross-reference such as <c>&lt;&lt;target,label&gt;&gt;</c>.</summary>
public sealed class AsciiDocCrossReferenceInline : AsciiDocInline {
    private string _target;
    private string? _text;

    internal AsciiDocCrossReferenceInline(AsciiDocSyntaxNode syntax, string target, string? text) : base(syntax) {
        _target = target;
        _text = text;
    }

    /// <summary>Anchor or document target.</summary>
    public string Target {
        get => _target;
        set { string normalized = value ?? string.Empty; AsciiDocText.EnsureSingleLine(normalized, nameof(value)); SetValue(ref _target, normalized); }
    }

    /// <summary>Optional visible reference text.</summary>
    public string? Text {
        get => _text;
        set { if (value != null) AsciiDocText.EnsureSingleLine(value, nameof(value)); SetValue(ref _text, value); }
    }

    internal override string WriteCore(AsciiDocWriterContext context) =>
        "<<" + Target + (Text == null ? string.Empty : "," + Text) + ">>";
}

/// <summary>Inline anchor such as <c>[[id,reference text]]</c>.</summary>
public sealed class AsciiDocAnchorInline : AsciiDocInline {
    private string _id;
    private string? _referenceText;

    internal AsciiDocAnchorInline(AsciiDocSyntaxNode syntax, string id, string? referenceText) : base(syntax) {
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

    internal override string WriteCore(AsciiDocWriterContext context) =>
        "[[" + Id + (ReferenceText == null ? string.Empty : "," + ReferenceText) + "]]";
}
