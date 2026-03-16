using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using Omd = OfficeIMO.Markdown;

namespace OfficeIMO.Word.Markdown {
    internal partial class MarkdownToWordConverter {
        private readonly struct InlineFormatState {
            public InlineFormatState(
                bool bold,
                bool italic,
                bool strike,
                UnderlineValues? underline,
                HighlightColorValues? highlight,
                VerticalPositionValues? verticalTextAlignment) {
                Bold = bold;
                Italic = italic;
                Strike = strike;
                Underline = underline;
                Highlight = highlight;
                VerticalTextAlignment = verticalTextAlignment;
            }

            public bool Bold { get; }
            public bool Italic { get; }
            public bool Strike { get; }
            public UnderlineValues? Underline { get; }
            public HighlightColorValues? Highlight { get; }
            public VerticalPositionValues? VerticalTextAlignment { get; }

            public InlineFormatState WithBold() => new InlineFormatState(bold: true, italic: Italic, strike: Strike, underline: Underline, highlight: Highlight, verticalTextAlignment: VerticalTextAlignment);
            public InlineFormatState WithItalic() => new InlineFormatState(bold: Bold, italic: true, strike: Strike, underline: Underline, highlight: Highlight, verticalTextAlignment: VerticalTextAlignment);
            public InlineFormatState WithStrike() => new InlineFormatState(bold: Bold, italic: Italic, strike: true, underline: Underline, highlight: Highlight, verticalTextAlignment: VerticalTextAlignment);
            public InlineFormatState WithUnderline(UnderlineValues underline) => new InlineFormatState(bold: Bold, italic: Italic, strike: Strike, underline: underline, highlight: Highlight, verticalTextAlignment: VerticalTextAlignment);
            public InlineFormatState WithHighlight(HighlightColorValues highlight) => new InlineFormatState(bold: Bold, italic: Italic, strike: Strike, underline: Underline, highlight: highlight, verticalTextAlignment: VerticalTextAlignment);
            public InlineFormatState WithVerticalTextAlignment(VerticalPositionValues verticalTextAlignment) => new InlineFormatState(bold: Bold, italic: Italic, strike: Strike, underline: Underline, highlight: Highlight, verticalTextAlignment: verticalTextAlignment);
        }

        private static WordParagraph AddRun(WordParagraph paragraph, string? text, InlineFormatState fmt, string? defaultFont) {
            var run = paragraph.AddText(text ?? string.Empty);
            if (fmt.Bold) run.SetBold();
            if (fmt.Italic) run.SetItalic();
            if (fmt.Underline.HasValue && fmt.Underline.Value != UnderlineValues.None) run.SetUnderline(fmt.Underline.Value);
            if (fmt.Strike) run.SetStrike();
            if (fmt.Highlight.HasValue && fmt.Highlight.Value != HighlightColorValues.None) run.SetHighlight(fmt.Highlight.Value);
            if (fmt.VerticalTextAlignment.HasValue) run.SetVerticalTextAlignment(fmt.VerticalTextAlignment.Value);
            if (!string.IsNullOrEmpty(defaultFont)) run.SetFontFamily(defaultFont!);
            return run;
        }

        private static WordParagraph CreateDetachedRun(WordDocument document, string? text, InlineFormatState fmt, string? defaultFont, bool forceMonospace = false) {
            var paragraph = new Paragraph();
            var run = new Run(new Text(text ?? string.Empty) {
                Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve
            });
            paragraph.Append(run);

            var wrapper = new WordParagraph(document, paragraph, run);
            if (fmt.Bold) wrapper.SetBold();
            if (fmt.Italic) wrapper.SetItalic();
            if (fmt.Underline.HasValue && fmt.Underline.Value != UnderlineValues.None) wrapper.SetUnderline(fmt.Underline.Value);
            if (fmt.Strike) wrapper.SetStrike();
            if (fmt.Highlight.HasValue && fmt.Highlight.Value != HighlightColorValues.None) wrapper.SetHighlight(fmt.Highlight.Value);
            if (fmt.VerticalTextAlignment.HasValue) wrapper.SetVerticalTextAlignment(fmt.VerticalTextAlignment.Value);

            if (forceMonospace) {
                wrapper.SetFontFamily(FontResolver.Resolve("monospace") ?? "Consolas");
            } else if (!string.IsNullOrEmpty(defaultFont)) {
                wrapper.SetFontFamily(defaultFont!);
            }

            return wrapper;
        }

        private static List<WordParagraph> BuildLinkLabelRunsOmd(
            IEnumerable<object> nodes,
            WordDocument document,
            InlineFormatState fmt,
            string? defaultFont) {
            var runs = new List<WordParagraph>();
            AppendLinkLabelRunsOmd(nodes, document, runs, fmt, defaultFont);
            return runs;
        }

        private static void AppendLinkLabelRunsOmd(
            IEnumerable<object> nodes,
            WordDocument document,
            List<WordParagraph> runs,
            InlineFormatState fmt,
            string? defaultFont) {
            foreach (var node in nodes) {
                switch (node) {
                    case null:
                        break;
                    case Omd.TextRun t:
                        runs.Add(CreateDetachedRun(document, t.Text, fmt, defaultFont));
                        break;
                    case Omd.HardBreakInline:
                        runs.Add(CreateDetachedRun(document, " ", fmt, defaultFont));
                        break;
                    case Omd.CodeSpanInline cs:
                        runs.Add(CreateDetachedRun(document, cs.Text, fmt, defaultFont, forceMonospace: true));
                        break;
                    case Omd.LinkInline l:
                        runs.Add(CreateDetachedRun(document, l.Text, fmt, defaultFont));
                        break;
                    case Omd.ImageInline im:
                        runs.Add(CreateDetachedRun(document, im.Alt ?? string.Empty, fmt, defaultFont));
                        break;
                    case Omd.ImageLinkInline il:
                        runs.Add(CreateDetachedRun(document, il.Alt ?? il.ImageUrl ?? il.LinkUrl ?? string.Empty, fmt, defaultFont));
                        break;
                    case Omd.FootnoteRefInline fn:
                        runs.Add(CreateDetachedRun(document, "[^" + fn.Label + "]", fmt, defaultFont));
                        break;
                    case Omd.BoldInline b:
                        runs.Add(CreateDetachedRun(document, b.Text, fmt.WithBold(), defaultFont));
                        break;
                    case Omd.ItalicInline it:
                        runs.Add(CreateDetachedRun(document, it.Text, fmt.WithItalic(), defaultFont));
                        break;
                    case Omd.BoldItalicInline bi:
                        runs.Add(CreateDetachedRun(document, bi.Text, fmt.WithBold().WithItalic(), defaultFont));
                        break;
                    case Omd.StrikethroughInline st:
                        runs.Add(CreateDetachedRun(document, st.Text, fmt.WithStrike(), defaultFont));
                        break;
                    case Omd.HighlightInline hi:
                        runs.Add(CreateDetachedRun(document, hi.Text, fmt.WithHighlight(HighlightColorValues.Yellow), defaultFont));
                        break;
                    case Omd.UnderlineInline un:
                        runs.Add(CreateDetachedRun(document, un.Text, fmt.WithUnderline(UnderlineValues.Single), defaultFont));
                        break;
                    case Omd.BoldSequenceInline bs:
                        AppendLinkLabelRunsOmd(bs.Inlines.Items ?? Array.Empty<object>(), document, runs, fmt.WithBold(), defaultFont);
                        break;
                    case Omd.ItalicSequenceInline iseq:
                        AppendLinkLabelRunsOmd(iseq.Inlines.Items ?? Array.Empty<object>(), document, runs, fmt.WithItalic(), defaultFont);
                        break;
                    case Omd.BoldItalicSequenceInline bis:
                        AppendLinkLabelRunsOmd(bis.Inlines.Items ?? Array.Empty<object>(), document, runs, fmt.WithBold().WithItalic(), defaultFont);
                        break;
                    case Omd.StrikethroughSequenceInline sts:
                        AppendLinkLabelRunsOmd(sts.Inlines.Items ?? Array.Empty<object>(), document, runs, fmt.WithStrike(), defaultFont);
                        break;
                    case Omd.HighlightSequenceInline hs:
                        AppendLinkLabelRunsOmd(hs.Inlines.Items ?? Array.Empty<object>(), document, runs, fmt.WithHighlight(HighlightColorValues.Yellow), defaultFont);
                        break;
                    case Omd.HtmlTagSequenceInline htmlTag:
                        AppendHtmlTagSequenceRunsOmd(htmlTag, document, runs, fmt, defaultFont);
                        break;
                    case Omd.HtmlRawInline htmlRaw:
                        if (!string.IsNullOrEmpty(htmlRaw.Html)) {
                            runs.Add(CreateDetachedRun(document, htmlRaw.Html, fmt, defaultFont));
                        }
                        break;
                    default:
                        break;
                }
            }
        }

        private static void AppendHtmlTagSequenceRunsOmd(
            Omd.HtmlTagSequenceInline htmlTag,
            WordDocument document,
            List<WordParagraph> runs,
            InlineFormatState fmt,
            string? defaultFont) {
            switch (htmlTag.TagName) {
                case "u":
                case "ins":
                    AppendLinkLabelRunsOmd(htmlTag.Inlines.Items ?? Array.Empty<object>(), document, runs, fmt.WithUnderline(UnderlineValues.Single), defaultFont);
                    break;
                case "sup":
                    AppendLinkLabelRunsOmd(htmlTag.Inlines.Items ?? Array.Empty<object>(), document, runs, fmt.WithVerticalTextAlignment(VerticalPositionValues.Superscript), defaultFont);
                    break;
                case "sub":
                    AppendLinkLabelRunsOmd(htmlTag.Inlines.Items ?? Array.Empty<object>(), document, runs, fmt.WithVerticalTextAlignment(VerticalPositionValues.Subscript), defaultFont);
                    break;
                case "q":
                    runs.Add(CreateDetachedRun(document, "\"", fmt, defaultFont));
                    AppendLinkLabelRunsOmd(htmlTag.Inlines.Items ?? Array.Empty<object>(), document, runs, fmt, defaultFont);
                    runs.Add(CreateDetachedRun(document, "\"", fmt, defaultFont));
                    break;
                default:
                    AppendLinkLabelRunsOmd(htmlTag.Inlines.Items ?? Array.Empty<object>(), document, runs, fmt, defaultFont);
                    break;
            }
        }

        private static InlineFormatState CreateDefaultInlineFormatState() =>
            new InlineFormatState(
                bold: false,
                italic: false,
                strike: false,
                underline: null,
                highlight: null,
                verticalTextAlignment: null);

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

            string? defaultFont = ResolveDefaultFontFamily(options);
            var list = inlines.Items ?? Array.Empty<object>();

            ProcessInlineNodesOmd(
                nodes: list,
                paragraph: paragraph,
                options: options,
                document: document,
                footnoteDefs: footnoteDefs,
                fmt: CreateDefaultInlineFormatState(),
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
                            var run = AddRun(paragraph, cs.Text, fmt, defaultFont);
                            var mono = FontResolver.Resolve("monospace") ?? "Consolas";
                            run.SetFontFamily(mono);
                            break;
                        }
                    case Omd.LinkInline l: {
                            try {
                                var uri = new Uri(l.Url, UriKind.RelativeOrAbsolute);
                                if (l.LabelInlines != null && (l.LabelInlines.Items?.Count ?? 0) > 0) {
                                    var labelRuns = BuildLinkLabelRunsOmd(l.LabelInlines.Items ?? Array.Empty<object>(), document, fmt, defaultFont);
                                    if (labelRuns.Count > 0) {
                                        WordHyperLink.AddHyperLink(paragraph, labelRuns, uri);
                                        break;
                                    }
                                }

                                var hl = paragraph.AddHyperLink(l.Text, uri);
                                ApplyHyperlinkFormattingOmd(hl, fmt, defaultFont);
                            } catch (UriFormatException ex) {
                                options.OnWarning?.Invoke($"Invalid URI '{l.Url}' - emitting as text. {ex.Message}");
                                if (l.LabelInlines != null && (l.LabelInlines.Items?.Count ?? 0) > 0) {
                                    ProcessInlineNodesOmd(l.LabelInlines.Items ?? Array.Empty<object>(), paragraph, options, document, footnoteDefs, fmt, defaultFont);
                                } else {
                                    AddRun(paragraph, l.Text, fmt, defaultFont);
                                }
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
                                ApplyHyperlinkFormattingOmd(hli, fmt, defaultFont);
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
                    case Omd.HighlightInline hi:
                        AddRun(paragraph, hi.Text, fmt.WithHighlight(HighlightColorValues.Yellow), defaultFont);
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
                    case Omd.HighlightSequenceInline hs:
                        ProcessInlineNodesOmd(hs.Inlines.Items ?? Array.Empty<object>(), paragraph, options, document, footnoteDefs, fmt.WithHighlight(HighlightColorValues.Yellow), defaultFont);
                        break;
                    case Omd.HtmlTagSequenceInline htmlTag:
                        ProcessHtmlTagSequenceInlineOmd(htmlTag, paragraph, options, document, footnoteDefs, fmt, defaultFont);
                        break;
                    case Omd.HtmlRawInline htmlRaw:
                        if (!string.IsNullOrEmpty(htmlRaw.Html)) {
                            AddRun(paragraph, htmlRaw.Html, fmt, defaultFont);
                        }
                        break;

                    default:
                        // Fallback: do not leak type names into the document.
                        break;
                }
            }
        }

        private static void ProcessHtmlTagSequenceInlineOmd(
            Omd.HtmlTagSequenceInline htmlTag,
            WordParagraph paragraph,
            MarkdownToWordOptions options,
            WordDocument document,
            IReadOnlyDictionary<string, string>? footnoteDefs,
            InlineFormatState fmt,
            string? defaultFont) {
            switch (htmlTag.TagName) {
                case "u":
                case "ins":
                    ProcessInlineNodesOmd(htmlTag.Inlines.Items ?? Array.Empty<object>(), paragraph, options, document, footnoteDefs, fmt.WithUnderline(UnderlineValues.Single), defaultFont);
                    break;
                case "sup":
                    ProcessInlineNodesOmd(htmlTag.Inlines.Items ?? Array.Empty<object>(), paragraph, options, document, footnoteDefs, fmt.WithVerticalTextAlignment(VerticalPositionValues.Superscript), defaultFont);
                    break;
                case "sub":
                    ProcessInlineNodesOmd(htmlTag.Inlines.Items ?? Array.Empty<object>(), paragraph, options, document, footnoteDefs, fmt.WithVerticalTextAlignment(VerticalPositionValues.Subscript), defaultFont);
                    break;
                case "q":
                    AddRun(paragraph, "\"", fmt, defaultFont);
                    ProcessInlineNodesOmd(htmlTag.Inlines.Items ?? Array.Empty<object>(), paragraph, options, document, footnoteDefs, fmt, defaultFont);
                    AddRun(paragraph, "\"", fmt, defaultFont);
                    break;
                default:
                    ProcessInlineNodesOmd(htmlTag.Inlines.Items ?? Array.Empty<object>(), paragraph, options, document, footnoteDefs, fmt, defaultFont);
                    break;
            }
        }

        private static void ApplyHyperlinkFormattingOmd(WordParagraph hyperlink, InlineFormatState fmt, string? defaultFont) {
            if (fmt.Bold) hyperlink.SetBold();
            if (fmt.Italic) hyperlink.SetItalic();
            if (fmt.Underline.HasValue && fmt.Underline.Value != UnderlineValues.None) hyperlink.SetUnderline(fmt.Underline.Value);
            if (fmt.Strike) hyperlink.SetStrike();
            if (fmt.Highlight.HasValue && fmt.Highlight.Value != HighlightColorValues.None) hyperlink.SetHighlight(fmt.Highlight.Value);
            if (fmt.VerticalTextAlignment.HasValue) hyperlink.SetVerticalTextAlignment(fmt.VerticalTextAlignment.Value);
            if (!string.IsNullOrEmpty(defaultFont)) hyperlink.SetFontFamily(defaultFont!);
        }
    }
}
