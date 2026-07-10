namespace OfficeIMO.AsciiDoc;

/// <summary>Base class for a source-backed AsciiDoc inline node.</summary>
public abstract class AsciiDocInline {
    private bool _isModified;

    internal AsciiDocInline(AsciiDocSyntaxNode syntax) {
        Syntax = syntax ?? throw new ArgumentNullException(nameof(syntax));
    }

    /// <summary>Lossless syntax backing this node.</summary>
    public AsciiDocSyntaxNode Syntax { get; }

    /// <summary>Exact original source span.</summary>
    public AsciiDocSourceSpan Span => Syntax.Span;

    /// <summary>Exact original source characters.</summary>
    public string OriginalText => Syntax.OriginalText;

    /// <summary>True when this node or a nested node has changed.</summary>
    public virtual bool IsModified => _isModified;

    internal string Write(AsciiDocWriterContext context) {
        if (context.Mode == AsciiDocWriterMode.Preserve && !IsModified) return OriginalText;
        return WriteCore(context);
    }

    internal abstract string WriteCore(AsciiDocWriterContext context);

    internal bool SetValue<T>(ref T field, T value) {
        if (EqualityComparer<T>.Default.Equals(field, value)) return false;
        field = value;
        _isModified = true;
        return true;
    }
}

/// <summary>Ordered inline content within a heading, paragraph, list item, or formatted span.</summary>
public sealed class AsciiDocInlineSequence {
    private readonly IReadOnlyList<AsciiDocInline> _items;

    internal AsciiDocInlineSequence(AsciiDocSyntaxNode syntax, IReadOnlyList<AsciiDocInline> items) {
        Syntax = syntax;
        _items = items;
    }

    /// <summary>Lossless syntax for the whole sequence.</summary>
    public AsciiDocSyntaxNode Syntax { get; }

    /// <summary>Inline nodes in source order.</summary>
    public IReadOnlyList<AsciiDocInline> Items => _items;

    /// <summary>True when any inline node has changed.</summary>
    public bool IsModified => Items.Any(static item => item.IsModified);

    /// <summary>Writes current inline content in preserve mode.</summary>
    public string ToAsciiDoc() => Write(new AsciiDocWriterContext(AsciiDocWriterMode.Preserve, "\n"));

    internal string Write(AsciiDocWriterContext context) {
        if (context.Mode == AsciiDocWriterMode.Preserve && !IsModified) return Syntax.OriginalText;
        var builder = new StringBuilder(Syntax.OriginalText.Length);
        for (int index = 0; index < Items.Count; index++) builder.Append(Items[index].Write(context));
        return builder.ToString();
    }
}

/// <summary>Literal inline text not assigned another semantic kind.</summary>
public sealed class AsciiDocTextInline : AsciiDocInline {
    private string _text;

    internal AsciiDocTextInline(AsciiDocSyntaxNode syntax, string text) : base(syntax) {
        _text = text;
    }

    /// <summary>Current text.</summary>
    public string Text {
        get => _text;
        set => SetValue(ref _text, value ?? string.Empty);
    }

    internal override string WriteCore(AsciiDocWriterContext context) =>
        context.Mode == AsciiDocWriterMode.Canonical ? AsciiDocText.NormalizeLineEndings(Text, context.LineEnding) : Text;
}
