using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO;

namespace OfficeIMO.Reader.DocBook;

internal static partial class DocBookReaderAdapter {
    private const string XLinkHrefName = "{http://www.w3.org/1999/xlink}href";

    private static bool TryBuildInlineFragments(
        OfficeDocumentModelNode node,
        out IReadOnlyList<InlineFragment> fragments) {
        var result = new List<InlineFragment>();
        bool hasTarget = false;
        if (string.Equals(node.Kind, "link", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(node.Kind, "cross-reference", StringComparison.OrdinalIgnoreCase)) {
            string? destination = GetInlineDestination(node);
            if (!string.IsNullOrEmpty(destination)) {
                string label = string.IsNullOrEmpty(node.Text) ? GetInlinePlainText(node) : node.Text;
                if (label.Length > 0 || string.Equals(node.Kind, "cross-reference", StringComparison.OrdinalIgnoreCase)) {
                    if (label.Length == 0) label = destination![0] == '#' ? destination.Substring(1) : destination;
                    result.Add(new InlineFragment(label, destination));
                    fragments = result;
                    return true;
                }
            }
        }
        foreach (OfficeDocumentModelNode child in node.Children) Append(child, string.Empty, string.Empty);
        fragments = result;
        return hasTarget && result.Count > 0;

        void Append(OfficeDocumentModelNode child, string prefix, string suffix) {
            if (IsIndexTerm(child.Kind)) return;
            bool emphasis = child.Kind == "extension:emphasis" || child.Kind == "extension:{http://docbook.org/ns/docbook}emphasis";
            bool literal = child.Kind == "extension:literal" || child.Kind == "extension:{http://docbook.org/ns/docbook}literal";
            if (literal) {
                string text = GetInlinePlainText(child);
                if (text.Length == 0) return;
                int longestRun = 0, run = 0;
                foreach (char character in text) { run = character == '`' ? run + 1 : 0; longestRun = Math.Max(longestRun, run); }
                string fence = new string('`', longestRun + 1);
                bool padding = text.StartsWith("`", StringComparison.Ordinal) || text.EndsWith("`", StringComparison.Ordinal) ||
                    (text.StartsWith(" ", StringComparison.Ordinal) && text.EndsWith(" ", StringComparison.Ordinal) && text.Trim(' ').Length > 0);
                string pad = padding ? " " : string.Empty;
                result.Add(new InlineFragment(text, null, prefix + fence + pad, pad + fence + suffix, false));
                hasTarget = true;
                return;
            }
            if (emphasis) {
                bool strong = child.Attributes.TryGetValue("role", out string? role) && (role == "bold" || role == "strong");
                string text = GetInlinePlainText(child);
                bool boundaryWhitespace = text.Length > 0 && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1]));
                // Links and nested styles split this element into independently wrapped fragments.
                // Markdown delimiters around those fragments may touch whitespace even when the
                // complete emphasis element does not, so use HTML wrappers for structured content.
                bool useInlineHtml = boundaryWhitespace || child.Children.Any(item => !string.Equals(item.Kind, "text", StringComparison.OrdinalIgnoreCase));
                string opening = useInlineHtml ? (strong ? "<strong>" : "<em>") : (strong ? "**" : "*");
                string closing = useInlineHtml ? (strong ? "</strong>" : "</em>") : opening;
                prefix += opening;
                suffix = closing + suffix;
                hasTarget = true;
            }
            if (string.Equals(child.Kind, "link", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(child.Kind, "cross-reference", StringComparison.OrdinalIgnoreCase)) {
                string? destination = GetInlineDestination(child);
                string label = string.IsNullOrEmpty(child.Text) ? GetInlinePlainText(child) : child.Text;
                if (!string.IsNullOrEmpty(destination)) {
                    if (label.Length > 0 || string.Equals(child.Kind, "cross-reference", StringComparison.OrdinalIgnoreCase)) {
                        if (label.Length == 0) label = destination![0] == '#' ? destination.Substring(1) : destination;
                        result.Add(new InlineFragment(label, destination, prefix, suffix));
                        hasTarget = true;
                        return;
                    }
                }
            }

            if (string.Equals(child.Kind, "text", StringComparison.OrdinalIgnoreCase)) {
                AddPlain(child.Text, prefix, suffix);
                return;
            }
            if (child.Children.Count > 0) {
                foreach (OfficeDocumentModelNode grandchild in child.Children) Append(grandchild, prefix, suffix);
            } else {
                AddPlain(child.Text, prefix, suffix);
            }
        }

        void AddPlain(string text, string prefix, string suffix) {
            if (text.Length == 0) return;
            if (result.Count > 0 && result[result.Count - 1].Destination == null &&
                result[result.Count - 1].StylePrefix == prefix && result[result.Count - 1].StyleSuffix == suffix && result[result.Count - 1].EscapesMarkdownText) {
                InlineFragment previous = result[result.Count - 1];
                result[result.Count - 1] = new InlineFragment(previous.Text + text, null, prefix, suffix);
            } else {
                result.Add(new InlineFragment(text, null, prefix, suffix));
            }
        }
    }

    private static string? GetInlineDestination(OfficeDocumentModelNode node) {
        if (node.Attributes.TryGetValue(XLinkHrefName, out string? href) && !string.IsNullOrWhiteSpace(href)) return href;
        if (node.Attributes.TryGetValue("url", out string? url) && !string.IsNullOrWhiteSpace(url)) return url;
        if (node.Attributes.TryGetValue("linkend", out string? linkEnd) && !string.IsNullOrWhiteSpace(linkEnd)) return "#" + linkEnd;
        return null;
    }

    private static string GetInlinePlainText(OfficeDocumentModelNode node) {
        if (!string.IsNullOrEmpty(node.Text)) return node.Text;
        var parts = new List<string>();
        Add(node);
        return string.Concat(parts);

        void Add(OfficeDocumentModelNode current) {
            if (IsIndexTerm(current.Kind)) return;
            if (string.Equals(current.Kind, "text", StringComparison.OrdinalIgnoreCase) && current.Text.Length > 0) {
                parts.Add(current.Text);
                return;
            }
            foreach (OfficeDocumentModelNode child in current.Children) Add(child);
        }
    }

    private sealed class InlineFragment {
        internal InlineFragment(string text, string? destination, string prefix = "", string suffix = "", bool escapeText = true) {
            Text = text;
            Destination = destination;
            StylePrefix = prefix;
            StyleSuffix = suffix;
            EscapesMarkdownText = escapeText;
        }

        internal string Text { get; }
        internal string? Destination { get; }
        internal string StylePrefix { get; }
        internal string StyleSuffix { get; }
        internal string MarkdownPrefix => (Destination == null ? string.Empty : "[") + StylePrefix;
        internal string MarkdownSuffix => StyleSuffix + (Destination == null ? string.Empty : "](" + EscapeDestination(Destination) + ")");
        internal bool EscapesMarkdownText { get; }

        private static string EscapeDestination(string value) {
            var escaped = new System.Text.StringBuilder(value.Length);
            foreach (char character in value) {
                if (char.IsWhiteSpace(character)) {
                    foreach (byte utf8Byte in System.Text.Encoding.UTF8.GetBytes(character.ToString())) {
                        escaped.Append('%').Append(utf8Byte.ToString("X2"));
                    }
                } else if (character == '\\' || character == '(' || character == ')') {
                    escaped.Append('\\').Append(character);
                } else {
                    escaped.Append(character);
                }
            }
            return escaped.ToString();
        }
    }
}
