using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Markdown;

/// <summary>
/// Unordered list supporting plain items and task (checklist) items.
/// </summary>
public sealed class UnorderedListBlock : IMarkdownBlock {
    /// <summary>List items.</summary>
    public List<ListItem> Items { get; } = new List<ListItem>();
    /// <inheritdoc />
    public string RenderMarkdown() => string.Join("\n", Items.Select(i => i.ToMarkdownListLine()));
    /// <inheritdoc />
    public string RenderHtml() => "<ul>" + string.Concat(Items.Select(i => i.ToHtmlListItem())) + "</ul>";
}
