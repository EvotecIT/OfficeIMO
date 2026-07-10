namespace OfficeIMO.AsciiDoc;

/// <summary>Source-backed list continuation marker that binds a following block to an item.</summary>
public sealed class AsciiDocListContinuation : AsciiDocBlock {
    internal AsciiDocListContinuation(AsciiDocSyntaxNode syntax, string trailingLineEnding) : base(syntax, trailingLineEnding) { }

    /// <summary>List item receiving the attached block.</summary>
    public AsciiDocListItem? TargetItem { get; internal set; }

    /// <summary>Block attached by this marker.</summary>
    public AsciiDocBlock? AttachedBlock { get; internal set; }

    internal override string WriteCore(AsciiDocWriterContext context) => "+" + EffectiveTrailingLineEnding(context);
}
