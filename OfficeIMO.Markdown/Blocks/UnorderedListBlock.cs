using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Markdown;

public sealed class UnorderedListBlock : IMarkdownBlock {
    public List<ListItem> Items { get; } = new List<ListItem>();
    public string RenderMarkdown() => string.Join("\n", Items.Select(i => "- " + i.RenderMarkdown()));
    public string RenderHtml() => "<ul>" + string.Concat(Items.Select(i => "<li>" + i.RenderHtml() + "</li>")) + "</ul>";
}

