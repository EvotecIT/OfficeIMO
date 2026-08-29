using System;
using System.Collections.Generic;
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
        foreach (OfficeDocumentModelNode child in node.Children) Append(child);
        fragments = result;
        return hasTarget && result.Count > 0;

        void Append(OfficeDocumentModelNode child) {
            if (string.Equals(child.Kind, "link", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(child.Kind, "cross-reference", StringComparison.OrdinalIgnoreCase)) {
                string? destination = GetInlineDestination(child);
                string label = string.IsNullOrEmpty(child.Text) ? GetInlinePlainText(child) : child.Text;
                if (!string.IsNullOrEmpty(destination)) {
                    if (label.Length > 0 || string.Equals(child.Kind, "cross-reference", StringComparison.OrdinalIgnoreCase)) {
                        if (label.Length == 0) label = destination![0] == '#' ? destination.Substring(1) : destination;
                        result.Add(new InlineFragment(label, destination));
                        hasTarget = true;
                        return;
                    }
                }
            }

            if (string.Equals(child.Kind, "text", StringComparison.OrdinalIgnoreCase)) {
                AddPlain(child.Text);
                return;
            }
            if (child.Children.Count > 0) {
                foreach (OfficeDocumentModelNode grandchild in child.Children) Append(grandchild);
            } else {
                AddPlain(child.Text);
            }
        }

        void AddPlain(string text) {
            if (text.Length == 0) return;
            if (result.Count > 0 && result[result.Count - 1].Destination == null) {
                InlineFragment previous = result[result.Count - 1];
                result[result.Count - 1] = new InlineFragment(previous.Text + text, null);
            } else {
                result.Add(new InlineFragment(text, null));
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
            if (string.Equals(current.Kind, "text", StringComparison.OrdinalIgnoreCase) && current.Text.Length > 0) {
                parts.Add(current.Text);
                return;
            }
            foreach (OfficeDocumentModelNode child in current.Children) Add(child);
        }
    }

    private sealed class InlineFragment {
        internal InlineFragment(string text, string? destination) {
            Text = text;
            Destination = destination;
        }

        internal string Text { get; }
        internal string? Destination { get; }

        internal string ToMarkdown(string text) {
            if (Destination == null) return text;
            return "[" + EscapeLabel(text) + "](" + EscapeDestination(Destination) + ")";
        }

        private static string EscapeLabel(string value) =>
            value.Replace("\\", "\\\\").Replace("[", "\\[").Replace("]", "\\]");

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
