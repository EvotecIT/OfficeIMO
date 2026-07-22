namespace OfficeIMO.Markdown;

/// <summary>
/// Identifies an inline node whose complete semantic value is literal text.
/// </summary>
public interface ILiteralTextMarkdownInline : IMarkdownInline {
    /// <summary>Literal semantic text represented by the inline node.</summary>
    string Text { get; }
}
