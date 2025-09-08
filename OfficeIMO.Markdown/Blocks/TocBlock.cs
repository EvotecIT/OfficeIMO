using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeIMO.Markdown;

/// <summary>
/// Table of contents block generated from document headings.
/// </summary>
public sealed class TocBlock : IMarkdownBlock {
    /// <summary>Single Table of Contents entry.</summary>
    public sealed class Entry {
        /// <summary>Heading level (1..6).</summary>
        public int Level { get; set; }
        /// <summary>Heading text.</summary>
        public string Text { get; set; } = string.Empty;
        /// <summary>Anchor id (slug) without leading '#'.</summary>
        public string Anchor { get; set; } = string.Empty;
    }

    /// <summary>When true, renders an ordered list; otherwise unordered.</summary>
    public bool Ordered { get; set; }
    /// <summary>Entries included in the TOC.</summary>
    public List<Entry> Entries { get; } = new List<Entry>();

    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() {
        if (Entries.Count == 0) return string.Empty;
        int baseLevel = Entries.Min(e => e.Level);
        var sb = new StringBuilder();
        foreach (var e in Entries) {
            int indent = e.Level - baseLevel;
            string pad = new string(' ', indent * 2);
            string marker = Ordered ? "1." : "-";
            sb.AppendLine($"{pad}{marker} [{e.Text}](#{e.Anchor})");
        }
        return sb.ToString().TrimEnd();
    }

    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() {
        if (Entries.Count == 0) return string.Empty;
        int baseLevel = Entries.Min(e => e.Level);
        var sb = new StringBuilder();
        // Build a nested list structure
        int currentLevel = baseLevel - 1;
        void OpenList(int level) { sb.Append(Ordered ? "<ol>" : "<ul>"); }
        void CloseList(int level) { sb.Append(Ordered ? "</ol>" : "</ul>"); }

        foreach (var e in Entries) {
            while (currentLevel < e.Level) { OpenList(++currentLevel); }
            while (currentLevel > e.Level) { CloseList(currentLevel--); }
            sb.Append($"<li><a href=\"#{System.Net.WebUtility.HtmlEncode(e.Anchor)}\">{System.Net.WebUtility.HtmlEncode(e.Text)}</a></li>");
        }
        while (currentLevel >= baseLevel) { CloseList(currentLevel--); }
        return sb.ToString();
    }
}
