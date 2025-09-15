using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
// Intentionally avoid heavy regex use; simple scanning is used for resilience and speed.

namespace OfficeIMO.Markdown;

/// <summary>
/// Parses Markdown text into OfficeIMO.Markdown's typed object model (<see cref="MarkdownDoc"/>, blocks, and inlines).
///
/// Scope: intentionally focused on the syntax that <see cref="MarkdownDoc"/> currently emits so we can
/// round‑trip what we generate. Reader behavior is controlled via <see cref="MarkdownReaderOptions"/>.
/// </summary>
public static partial class MarkdownReader {
    /// <summary>
    /// Parses Markdown text into a <see cref="MarkdownDoc"/> with typed blocks and basic inlines.
    /// </summary>
    public static MarkdownDoc Parse(string markdown, MarkdownReaderOptions? options = null) {
        options ??= new MarkdownReaderOptions();
        var doc = MarkdownDoc.Create();
        if (string.IsNullOrEmpty(markdown)) return doc;

        // Normalize line endings and split. Keep empty lines significant for block boundaries.
        var text = markdown.Replace("\r\n", "\n").Replace('\r', '\n');
        var lines = text.Split('\n');
        int i = 0;

        // Front matter (YAML) only if it's the very first thing in the file
        if (options.FrontMatter && i < lines.Length && lines[i].Trim() == "---") {
            int start = i + 1;
            int end = -1;
            for (int j = start; j < lines.Length; j++) { if (lines[j].Trim() == "---") { end = j; break; } }
            if (end > start) {
                var dict = ParseFrontMatter(lines, start, end - 1);
                if (dict.Count > 0) doc.Add(FrontMatterBlock.FromObject(dict));
                i = end + 1;
                // optional blank line after front matter
                if (i < lines.Length && string.IsNullOrWhiteSpace(lines[i])) i++;
            }
        }

        var pipeline = MarkdownReaderPipeline.Default();
        var state = new MarkdownReaderState();
        // Pre-scan for reference-style link definitions so inline refs in earlier paragraphs can resolve
        PreScanReferenceLinkDefinitions(lines, state);
        while (i < lines.Length) {
            if (string.IsNullOrWhiteSpace(lines[i])) { i++; continue; }
            bool matched = false;
            var parsers = pipeline.Parsers;
            for (int p = 0; p < parsers.Count; p++) {
                if (parsers[p].TryParse(lines, ref i, options, doc, state)) { matched = true; break; }
            }
            if (!matched) i++; // defensive: avoid infinite loop
        }

        return doc;
    }

    /// <summary>Parses a Markdown file path into a <see cref="MarkdownDoc"/>.</summary>
    public static MarkdownDoc ParseFile(string path, MarkdownReaderOptions? options = null) {
        string text = File.ReadAllText(path, Encoding.UTF8);
        return Parse(text, options);
    }

    private static void PreScanReferenceLinkDefinitions(string[] lines, MarkdownReaderState state) {
        for (int idx = 0; idx < lines.Length; idx++) {
            var line = lines[idx]; if (string.IsNullOrWhiteSpace(line)) continue;
            var t = line.Trim(); if (t.Length < 5 || t[0] != '[') continue;
            int rb = t.IndexOf(']'); if (rb <= 1) continue;
            if (rb + 1 >= t.Length || t[rb + 1] != ':') continue;
            string label = t.Substring(1, rb - 1);
            string rest = t.Substring(rb + 2).Trim(); if (string.IsNullOrEmpty(rest)) continue;
            string url = rest; string? title = null;
            bool urlWasBracketed = false;
            if (rest.Length > 0 && rest[0] == '<') {
                int gt = rest.IndexOf('>');
                if (gt > 1) {
                    url = rest.Substring(1, gt - 1);
                    rest = rest.Substring(gt + 1).Trim();
                    urlWasBracketed = true;
                }
            }
            int q = rest.IndexOf('"');
            if (q >= 0) {
                if (!urlWasBracketed) {
                    url = rest.Substring(0, q).Trim();
                }
                int q2 = rest.LastIndexOf('"');
                if (q2 > q) title = rest.Substring(q + 1, q2 - q - 1);
            }
            if (!string.IsNullOrEmpty(label) && !string.IsNullOrEmpty(url)) state.LinkRefs[label] = (url, title);
        }
    }
}
