namespace OfficeIMO.Markdown;

/// <summary>
/// Markdown heading (ATX) block, levels 1–6.
/// </summary>
public sealed class HeadingBlock : IMarkdownBlock {
    /// <summary>Heading level constrained to [1,6].</summary>
    public int Level { get; }
    /// <summary>Heading text.</summary>
    public string Text { get; }
    /// <summary>
    /// Creates a new heading block.
    /// </summary>
    /// <param name="level">Desired level; constrained to [1,6].</param>
    /// <param name="text">Heading text.</param>
    public HeadingBlock(int level, string text) {
        // Manual clamp to support netstandard2.0 where Math.Clamp may not exist.
        if (level < 1) level = 1; else if (level > 6) level = 6;
        Level = level;
        Text = text ?? string.Empty;
    }
    /// <inheritdoc />
    public string RenderMarkdown() => new string('#', Level) + " " + Text;
    /// <inheritdoc />
    public string RenderHtml() => $"<h{Level}>{System.Net.WebUtility.HtmlEncode(Text)}</h{Level}>";
}
