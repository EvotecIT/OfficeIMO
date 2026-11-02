namespace OfficeIMO.Markdown;

public sealed partial class MarkdownTableBuilder {
    internal static MarkdownColumnAlignment ToMarkdownAlignment(OfficeAlignment? alignment) => alignment switch {
        null => MarkdownColumnAlignment.None,
        OfficeAlignment.Left or OfficeAlignment.Start => MarkdownColumnAlignment.Left,
        OfficeAlignment.Center => MarkdownColumnAlignment.Center,
        OfficeAlignment.Right or OfficeAlignment.End => MarkdownColumnAlignment.Right,
        _ => throw new ArgumentOutOfRangeException(nameof(alignment), "Markdown pipe tables support left, center and right alignment.")
    };
}
