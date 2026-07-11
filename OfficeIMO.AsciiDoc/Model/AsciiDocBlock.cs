namespace OfficeIMO.AsciiDoc;

/// <summary>
/// Base class for typed top-level AsciiDoc blocks backed by exact original syntax.
/// </summary>
public abstract class AsciiDocBlock {
    private bool _isModified;
    private readonly List<AsciiDocBlockAttributeList> _attributeLists = new List<AsciiDocBlockAttributeList>();
    private AsciiDocBlockTitle? _blockTitle;
    private AsciiDocBlockAnchor? _blockAnchor;

    internal AsciiDocBlock(AsciiDocSyntaxNode syntax, string trailingLineEnding) {
        Syntax = syntax ?? throw new ArgumentNullException(nameof(syntax));
        TrailingLineEnding = trailingLineEnding ?? string.Empty;
    }

    /// <summary>Lossless syntax node associated with this semantic block.</summary>
    public AsciiDocSyntaxNode Syntax { get; }

    /// <summary>Exact original source span.</summary>
    public AsciiDocSourceSpan Span => Syntax.Span;

    /// <summary>Exact original source text.</summary>
    public string OriginalText => Syntax.OriginalText;

    /// <summary>True when this block or one of its editable children has changed.</summary>
    public virtual bool IsModified => _isModified;

    /// <summary>Source-backed element attribute lists bound to this block.</summary>
    public IReadOnlyList<AsciiDocBlockAttributeList> AttributeLists => _attributeLists;

    /// <summary>Optional source-backed block title.</summary>
    public AsciiDocBlockTitle? BlockTitle => _blockTitle;

    /// <summary>Optional source-backed block anchor.</summary>
    public AsciiDocBlockAnchor? BlockAnchor => _blockAnchor;

    /// <summary>Effective block style from the last bound attribute list.</summary>
    public string? Style {
        get {
            for (int index = AttributeLists.Count - 1; index >= 0; index--) {
                string? style = AttributeLists[index].Attributes.Style;
                if (style != null) return style;
            }
            return null;
        }
    }

    internal string TrailingLineEnding { get; }

    internal string Write(AsciiDocWriterContext context) {
        if (context.Mode == AsciiDocWriterMode.Preserve && !IsModified) return OriginalText;
        return WriteCore(context);
    }

    internal abstract string WriteCore(AsciiDocWriterContext context);

    internal string EffectiveTrailingLineEnding(AsciiDocWriterContext context) =>
        context.Mode == AsciiDocWriterMode.Preserve ? TrailingLineEnding : (TrailingLineEnding.Length == 0 ? string.Empty : context.LineEnding);

    internal bool SetValue<T>(ref T field, T value) {
        if (EqualityComparer<T>.Default.Equals(field, value)) return false;
        field = value;
        _isModified = true;
        return true;
    }

    internal void AddAttributeList(AsciiDocBlockAttributeList attributeList) {
        _attributeLists.Add(attributeList);
    }

    internal void SetBlockTitle(AsciiDocBlockTitle title) => _blockTitle = title;

    internal void SetBlockAnchor(AsciiDocBlockAnchor anchor) => _blockAnchor = anchor;
}
