using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using Omd = OfficeIMO.Markdown;

namespace OfficeIMO.Word.Markdown {
    internal partial class MarkdownToWordConverter {
        private readonly struct InlineFormatState {
            public InlineFormatState(bool bold, bool italic, bool strike, UnderlineValues? underline) {
                Bold = bold;
                Italic = italic;
                Strike = strike;
                Underline = underline;
            }

            public bool Bold { get; }
            public bool Italic { get; }
            public bool Strike { get; }
            public UnderlineValues? Underline { get; }

            public InlineFormatState WithBold() => new InlineFormatState(bold: true, italic: Italic, strike: Strike, underline: Underline);
            public InlineFormatState WithItalic() => new InlineFormatState(bold: Bold, italic: true, strike: Strike, underline: Underline);
            public InlineFormatState WithStrike() => new InlineFormatState(bold: Bold, italic: Italic, strike: true, underline: Underline);
            public InlineFormatState WithUnderline(UnderlineValues underline) => new InlineFormatState(bold: Bold, italic: Italic, strike: Strike, underline: underline);
        }

        private static WordParagraph AddRun(WordParagraph paragraph, string? text, InlineFormatState fmt, string? defaultFont) {
            var run = paragraph.AddText(text ?? string.Empty);
            if (fmt.Bold) run.SetBold();
            if (fmt.Italic) run.SetItalic();
            if (fmt.Underline.HasValue && fmt.Underline.Value != UnderlineValues.None) run.SetUnderline(fmt.Underline.Value);
            if (fmt.Strike) run.SetStrike();
            if (!string.IsNullOrEmpty(defaultFont)) run.SetFontFamily(defaultFont!);
            return run;
        }

        /// <summary>
        /// Processes OfficeIMO.Markdown inline sequence into Word runs.
        /// </summary>
        private static void ProcessInlinesOmd(
            Omd.InlineSequence inlines,
            WordParagraph paragraph,
            MarkdownToWordOptions options,
            WordDocument document,
            IReadOnlyDictionary<string, string>? footnoteDefs = null
        ) {
            if (inlines == null) return;

            string? defaultFont = options.FontFamily;
            var list = inlines.Items ?? Array.Empty<object>();

            ProcessInlineNodesOmd(
                nodes: list,
                paragraph: paragraph,
                options: options,
                document: document,
                footnoteDefs: footnoteDefs,
                fmt: new InlineFormatState(bold: false, italic: false, strike: false, underline: null),
                defaultFont: defaultFont
            );
        }

        private static void ProcessInlineNodesOmd(
            IEnumerable<object> nodes,
            WordParagraph paragraph,
            MarkdownToWordOptions options,
            WordDocument document,
            IReadOnlyDictionary<string, string>? footnoteDefs,
            InlineFormatState fmt,
            string? defaultFont
        ) {
            foreach (var node in nodes) {
                switch (node) {
                    case null:
                        break;
                    case Omd.TextRun t:
                        AddRun(paragraph, t.Text, fmt, defaultFont);
                        break;
                    case Omd.HardBreakInline:
                        paragraph.AddBreak();
                        break;
                    case Omd.CodeSpanInline cs: {
                            // Represent inline code using monospace font so Word -> Markdown can recover it.
                            var run = paragraph.AddText(cs.Text ?? string.Empty);
                            var mono = FontResolver.Resolve("monospace") ?? "Consolas";
                            run.SetFontFamily(mono);
                            break;
                        }
                    case Omd.LinkInline l: {
                            try {
                                var uri = new Uri(l.Url, UriKind.RelativeOrAbsolute);
                                var hl = paragraph.AddHyperLink(l.Text, uri);

                                // Best-effort: apply formatting to the hyperlink run.
                                if (fmt.Bold) hl.SetBold();
                                if (fmt.Italic) hl.SetItalic();
                                if (fmt.Underline.HasValue && fmt.Underline.Value != UnderlineValues.None) hl.SetUnderline(fmt.Underline.Value);
                                if (fmt.Strike) hl.SetStrike();
                                if (!string.IsNullOrEmpty(defaultFont)) hl.SetFontFamily(defaultFont!);
                            } catch (UriFormatException ex) {
                                options.OnWarning?.Invoke($"Invalid URI '{l.Url}' - emitting as text. {ex.Message}");
                                AddRun(paragraph, l.Text, fmt, defaultFont);
                            }
                            break;
                        }
                    case Omd.ImageLinkInline il: {
                            // Minimal mapping: insert hyperlink with alt text; inline image support is optional.
                            var linkUrl = il.LinkUrl ?? string.Empty;
                            var label = il.Alt ?? il.ImageUrl ?? linkUrl;
                            try {
                                if (string.IsNullOrEmpty(linkUrl)) {
                                    AddRun(paragraph, label, fmt, defaultFont);
                                    break;
                                }
                                var uri = new Uri(linkUrl, UriKind.RelativeOrAbsolute);
                                var hli = paragraph.AddHyperLink(label, uri);
                                if (fmt.Bold) hli.SetBold();
                                if (fmt.Italic) hli.SetItalic();
                                if (fmt.Underline.HasValue && fmt.Underline.Value != UnderlineValues.None) hli.SetUnderline(fmt.Underline.Value);
                                if (fmt.Strike) hli.SetStrike();
                                if (!string.IsNullOrEmpty(defaultFont)) hli.SetFontFamily(defaultFont!);
                            } catch (UriFormatException ex) {
                                options.OnWarning?.Invoke($"Invalid URI '{linkUrl}' - emitting alt text. {ex.Message}");
                                AddRun(paragraph, label, fmt, defaultFont);
                            }
                            break;
                        }
                    case Omd.ImageInline im:
                        // Inline images are not currently mapped; preserve alt text at least.
                        AddRun(paragraph, im.Alt ?? string.Empty, fmt, defaultFont);
                        break;
                    case Omd.FootnoteRefInline fn: {
                            string text = fn.Label;
                            if (footnoteDefs != null && footnoteDefs.TryGetValue(fn.Label, out var body)) {
                                text = body;
                            }
                            paragraph.AddFootNote(text);
                            break;
                        }

                    // Legacy builder-style inlines (flat text)
                    case Omd.BoldInline b:
                        AddRun(paragraph, b.Text, fmt.WithBold(), defaultFont);
                        break;
                    case Omd.ItalicInline it:
                        AddRun(paragraph, it.Text, fmt.WithItalic(), defaultFont);
                        break;
                    case Omd.BoldItalicInline bi:
                        AddRun(paragraph, bi.Text, fmt.WithBold().WithItalic(), defaultFont);
                        break;
                    case Omd.StrikethroughInline st:
                        AddRun(paragraph, st.Text, fmt.WithStrike(), defaultFont);
                        break;
                    case Omd.UnderlineInline un:
                        AddRun(paragraph, un.Text, fmt.WithUnderline(UnderlineValues.Single), defaultFont);
                        break;

                    // Reader-produced nested inlines
                    case Omd.BoldSequenceInline bs:
                        ProcessInlineNodesOmd(bs.Inlines.Items ?? Array.Empty<object>(), paragraph, options, document, footnoteDefs, fmt.WithBold(), defaultFont);
                        break;
                    case Omd.ItalicSequenceInline iseq:
                        ProcessInlineNodesOmd(iseq.Inlines.Items ?? Array.Empty<object>(), paragraph, options, document, footnoteDefs, fmt.WithItalic(), defaultFont);
                        break;
                    case Omd.BoldItalicSequenceInline bis:
                        ProcessInlineNodesOmd(bis.Inlines.Items ?? Array.Empty<object>(), paragraph, options, document, footnoteDefs, fmt.WithBold().WithItalic(), defaultFont);
                        break;
                    case Omd.StrikethroughSequenceInline sts:
                        ProcessInlineNodesOmd(sts.Inlines.Items ?? Array.Empty<object>(), paragraph, options, document, footnoteDefs, fmt.WithStrike(), defaultFont);
                        break;

                    default:
                        // Fallback: do not leak type names into the document.
                        break;
                }
            }
        }
    }
}
