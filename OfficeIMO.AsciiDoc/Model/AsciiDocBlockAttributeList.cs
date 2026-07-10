namespace OfficeIMO.AsciiDoc;

/// <summary>Source-preserving element attribute list bound to the following block.</summary>
public sealed class AsciiDocBlockAttributeList : AsciiDocBlock, IAsciiDocBlockMetadata {
    private string _content;
    private AsciiDocElementAttributes _attributes;

    internal AsciiDocBlockAttributeList(AsciiDocSyntaxNode syntax, string content, string trailingLineEnding)
        : base(syntax, trailingLineEnding) {
        _content = content;
        _attributes = AsciiDocAttributeListParser.Parse(content);
    }

    /// <summary>Content between square brackets.</summary>
    public string Content {
        get => _content;
        set {
            string normalized = value ?? string.Empty;
            AsciiDocText.EnsureSingleLine(normalized, nameof(value));
            if (SetValue(ref _content, normalized)) _attributes = AsciiDocAttributeListParser.Parse(normalized);
        }
    }

    /// <summary>Parsed positional, named, ID, role, and option entries.</summary>
    public AsciiDocElementAttributes Attributes => _attributes;

    /// <summary>Block this metadata line is bound to, or null when no block follows directly.</summary>
    public AsciiDocBlock? Target { get; internal set; }

    AsciiDocBlock? IAsciiDocBlockMetadata.Target { get => Target; set => Target = value; }

    internal override string WriteCore(AsciiDocWriterContext context) =>
        "[" + Content + "]" + EffectiveTrailingLineEnding(context);
}
