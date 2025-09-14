using OfficeIMO.Word;
using Omd = OfficeIMO.Markdown;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;

namespace OfficeIMO.Word.Markdown {
    internal partial class MarkdownToWordConverter {
        /// <summary>
        /// Processes OfficeIMO.Markdown inline sequence into Word runs.
        /// </summary>
        private static void ProcessInlinesOmd(Omd.InlineSequence inlines,
                                              WordParagraph paragraph,
                                              MarkdownToWordOptions options,
                                              WordDocument document,
                                              IReadOnlyDictionary<string, string>? footnoteDefs = null) {
            if (inlines == null) return;

            // InlineSequence stores a private list of objects; use rendered HTML to keep spacing
            // but we need typed access. We’ll reflect items via public Renderers: instead, we add
            // simple pattern-based handling by re-parsing segments when necessary.
            // Since InlineSequence is ours, we can read items by splitting RenderMarkdown conservatively
            // but a better approach is to expose items. For now, rely on known inline classes via dynamic.

            var field = typeof(Omd.InlineSequence).GetField("_inlines", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            var list = (System.Collections.Generic.List<object>) (field?.GetValue(inlines) ?? new System.Collections.Generic.List<object>());

            string? defaultFont = options.FontFamily;

            foreach (var node in list) {
                switch (node) {
                    case Omd.TextRun t:
                        var run = paragraph.AddText(t.Text);
                        if (!string.IsNullOrEmpty(defaultFont)) run.SetFontFamily(defaultFont!);
                        break;
                    case Omd.LinkInline l:
                        try {
                            var hl = paragraph.AddHyperLink(l.Text, new Uri(l.Url, UriKind.RelativeOrAbsolute));
                            if (!string.IsNullOrEmpty(defaultFont)) hl.SetFontFamily(defaultFont!);
                        } catch {
                            // Fallback to plain text if URI invalid
                            var r = paragraph.AddText(l.Text);
                            if (!string.IsNullOrEmpty(defaultFont)) r.SetFontFamily(defaultFont!);
                        }
                        break;
                    case Omd.BoldInline b:
                        var rb = paragraph.AddFormattedText(b.Text, bold: true);
                        if (!string.IsNullOrEmpty(defaultFont)) rb.SetFontFamily(defaultFont!);
                        break;
                    case Omd.BoldItalicInline bi:
                        var rbi = paragraph.AddFormattedText(bi.Text, bold: true, italic: true);
                        if (!string.IsNullOrEmpty(defaultFont)) rbi.SetFontFamily(defaultFont!);
                        break;
                    case Omd.ItalicInline it:
                        var ri = paragraph.AddFormattedText(it.Text, italic: true);
                        if (!string.IsNullOrEmpty(defaultFont)) ri.SetFontFamily(defaultFont!);
                        break;
                    case Omd.StrikethroughInline st:
                        var rs = paragraph.AddText(st.Text).SetStrike();
                        if (!string.IsNullOrEmpty(defaultFont)) rs.SetFontFamily(defaultFont!);
                        break;
                    case Omd.UnderlineInline un:
                        var ru = paragraph.AddText(un.Text).SetUnderline(UnderlineValues.Single);
                        if (!string.IsNullOrEmpty(defaultFont)) ru.SetFontFamily(defaultFont!);
                        break;
                    case Omd.CodeSpanInline cs:
                        var rc = paragraph.AddText(cs.Text);
                        var mono = FontResolver.Resolve("monospace") ?? "Consolas";
                        rc.SetFontFamily(mono);
                        break;
                    case Omd.ImageLinkInline il:
                        // Minimal mapping: insert hyperlink with alt text; images inside runs are supported but optional here.
                        try {
                            var hli = paragraph.AddHyperLink(il.Alt ?? il.ImageUrl ?? il.LinkUrl, new Uri(il.LinkUrl, UriKind.RelativeOrAbsolute));
                            if (!string.IsNullOrEmpty(defaultFont)) hli.SetFontFamily(defaultFont!);
                        } catch {
                            var ralt = paragraph.AddText(il.Alt ?? string.Empty);
                            if (!string.IsNullOrEmpty(defaultFont)) ralt.SetFontFamily(defaultFont!);
                        }
                        break;
                    case Omd.FootnoteRefInline fn:
                        string text = fn.Label;
                        if (footnoteDefs != null && footnoteDefs.TryGetValue(fn.Label, out var body)) text = body;
                        paragraph.AddFootNote(text);
                        break;
                    case Omd.HardBreakInline:
                        paragraph.AddBreak();
                        break;
                    default:
                        // Fallback: render markdown and insert as plain text
                        var str = node?.ToString();
                        if (!string.IsNullOrEmpty(str)) {
                            var r0 = paragraph.AddText(str!);
                            if (!string.IsNullOrEmpty(defaultFont)) r0.SetFontFamily(defaultFont!);
                        }
                        break;
                }
            }
        }
    }
}

