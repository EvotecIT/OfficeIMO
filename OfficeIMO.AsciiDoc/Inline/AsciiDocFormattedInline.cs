namespace OfficeIMO.AsciiDoc;

/// <summary>Formatting styles represented by AsciiDoc quote substitutions.</summary>
public enum AsciiDocInlineStyle {
    /// <summary>Bold text.</summary>
    Strong = 0,
    /// <summary>Italic text.</summary>
    Emphasis,
    /// <summary>Monospaced text.</summary>
    Monospace,
    /// <summary>Highlighted or role-styled text.</summary>
    Highlight,
    /// <summary>Subscript text.</summary>
    Subscript,
    /// <summary>Superscript text.</summary>
    Superscript
}

/// <summary>Source-backed constrained or unconstrained formatted phrase.</summary>
public sealed class AsciiDocFormattedInline : AsciiDocInline {
    internal AsciiDocFormattedInline(
        AsciiDocSyntaxNode syntax,
        AsciiDocInlineStyle style,
        string marker,
        AsciiDocInlineSequence content) : base(syntax) {
        Style = style;
        Marker = marker;
        Content = content;
    }

    /// <summary>Formatting semantic.</summary>
    public AsciiDocInlineStyle Style { get; }

    /// <summary>Original single or double formatting marker.</summary>
    public string Marker { get; }

    /// <summary>True when the double-marker unconstrained form was used.</summary>
    public bool IsUnconstrained => Marker.Length == 2;

    /// <summary>Nested inline content.</summary>
    public AsciiDocInlineSequence Content { get; }

    /// <inheritdoc />
    public override bool IsModified => base.IsModified || Content.IsModified;

    internal override string WriteCore(AsciiDocWriterContext context) =>
        Marker + Content.Write(context) + Marker;
}
