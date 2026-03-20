namespace OfficeIMO.Markdown;

/// <summary>
/// Plain-text extraction contract for inline nodes.
/// Implement this on custom inline types so headings, summaries, and other text-derived views can
/// safely read semantic text from the inline tree.
/// </summary>
public interface IPlainTextMarkdownInline {
    /// <summary>
    /// Appends the plain-text representation of the inline node to the supplied string builder.
    /// </summary>
    void AppendPlainText(System.Text.StringBuilder sb);
}
