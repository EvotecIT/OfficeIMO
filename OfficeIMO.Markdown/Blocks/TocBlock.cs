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
    /// <summary>Normalize indentation to the minimum included heading level (default true).</summary>
    public bool NormalizeLevels { get; set; } = true;
    /// <summary>Entries included in the TOC.</summary>
    public List<Entry> Entries { get; } = new List<Entry>();

    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() {
        if (Entries.Count == 0) return string.Empty;
        int baseLevel = NormalizeLevels ? Entries.Min(e => e.Level) : 1;
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
        int baseLevel = NormalizeLevels ? Entries.Min(e => e.Level) : 1;

        // Build a tree of nodes based on heading levels
        var root = new Node { Level = baseLevel - 1 };
        var stack = new System.Collections.Generic.Stack<Node>();
        stack.Push(root);
        foreach (var e in Entries) {
            // Pop until parent level is less than current entry level
            while (stack.Count > 0 && stack.Peek().Level >= e.Level) stack.Pop();
            var parent = stack.Peek();
            var node = new Node { Level = e.Level, Entry = e };
            parent.Children.Add(node);
            stack.Push(node);
        }

        var sb = new StringBuilder();
        RenderList(sb, root.Children, Ordered);
        return sb.ToString();

        static void RenderList(StringBuilder sb, System.Collections.Generic.IEnumerable<Node> nodes, bool ordered) {
            sb.Append(ordered ? "<ol>" : "<ul>");
            foreach (var n in nodes) {
                var e = n.Entry!;
                sb.Append("<li>");
                sb.Append($"<a href=\"#{System.Net.WebUtility.HtmlEncode(e.Anchor)}\">{System.Net.WebUtility.HtmlEncode(e.Text)}</a>");
                if (n.Children.Count > 0) RenderList(sb, n.Children, ordered);
                sb.Append("</li>");
            }
            sb.Append(ordered ? "</ol>" : "</ul>");
        }
    }

    private sealed class Node {
        public int Level { get; set; }
        public Entry? Entry { get; set; }
        public System.Collections.Generic.List<Node> Children { get; } = new System.Collections.Generic.List<Node>();
    }
}
