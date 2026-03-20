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
            Omd.InlineSequence? inlines,
            WordDocument document,
            InlineFormatState fmt,
            string? defaultFont) {
            var runs = new List<WordParagraph>();
            if (inlines == null) {
                return runs;
            }

            new DetachedInlineRunBuilder(document, runs, fmt, defaultFont).Visit(inlines);
            return runs;
        }

        private sealed class DetachedInlineRunBuilder : Omd.MarkdownVisitor {
            private readonly WordDocument _document;
            private readonly List<WordParagraph> _runs;
            private readonly InlineFormatState _fmt;
            private readonly string? _defaultFont;

            public DetachedInlineRunBuilder(
                WordDocument document,
                List<WordParagraph> runs,
                InlineFormatState fmt,
                string? defaultFont) {
                _document = document;
                _runs = runs;
                _fmt = fmt;
                _defaultFont = defaultFont;
            }

            private void VisitNested(Omd.MarkdownObject? node, InlineFormatState format) {
                if (node == null) {
                    return;
                }

                new DetachedInlineRunBuilder(_document, _runs, format, _defaultFont).Visit(node);
            }

            protected override void VisitTextRun(Omd.TextRun inline) =>
                _runs.Add(CreateDetachedRun(_document, inline.Text, _fmt, _defaultFont));

            protected override void VisitHardBreakInline(Omd.HardBreakInline inline) =>
                _runs.Add(CreateDetachedRun(_document, " ", _fmt, _defaultFont));

            protected override void VisitCodeSpanInline(Omd.CodeSpanInline inline) =>
                _runs.Add(CreateDetachedRun(_document, inline.Text, _fmt, _defaultFont, forceMonospace: true));

            protected override void VisitLinkInline(Omd.LinkInline inline) =>
                _runs.Add(CreateDetachedRun(_document, inline.Text, _fmt, _defaultFont));

            protected override void VisitImageInline(Omd.ImageInline inline) =>
                _runs.Add(CreateDetachedRun(_document, inline.Alt ?? string.Empty, _fmt, _defaultFont));

            protected override void VisitImageLinkInline(Omd.ImageLinkInline inline) =>
                _runs.Add(CreateDetachedRun(_document, inline.Alt ?? inline.ImageUrl ?? inline.LinkUrl ?? string.Empty, _fmt, _defaultFont));

            protected override void VisitFootnoteRefInline(Omd.FootnoteRefInline inline) =>
                _runs.Add(CreateDetachedRun(_document, "[^" + inline.Label + "]", _fmt, _defaultFont));

            protected override void VisitBoldInline(Omd.BoldInline inline) =>
                _runs.Add(CreateDetachedRun(_document, inline.Text, _fmt.WithBold(), _defaultFont));

            protected override void VisitItalicInline(Omd.ItalicInline inline) =>
                _runs.Add(CreateDetachedRun(_document, inline.Text, _fmt.WithItalic(), _defaultFont));

            protected override void VisitBoldItalicInline(Omd.BoldItalicInline inline) =>
                _runs.Add(CreateDetachedRun(_document, inline.Text, _fmt.WithBold().WithItalic(), _defaultFont));

            protected override void VisitStrikethroughInline(Omd.StrikethroughInline inline) =>
                _runs.Add(CreateDetachedRun(_document, inline.Text, _fmt.WithStrike(), _defaultFont));

            protected override void VisitHighlightInline(Omd.HighlightInline inline) =>
                _runs.Add(CreateDetachedRun(_document, inline.Text, _fmt.WithHighlight(HighlightColorValues.Yellow), _defaultFont));

            protected override void VisitUnderlineInline(Omd.UnderlineInline inline) =>
                _runs.Add(CreateDetachedRun(_document, inline.Text, _fmt.WithUnderline(UnderlineValues.Single), _defaultFont));

            protected override void VisitBoldSequenceInline(Omd.BoldSequenceInline inline) =>
                VisitNested(inline.Inlines, _fmt.WithBold());

            protected override void VisitItalicSequenceInline(Omd.ItalicSequenceInline inline) =>
                VisitNested(inline.Inlines, _fmt.WithItalic());

            protected override void VisitBoldItalicSequenceInline(Omd.BoldItalicSequenceInline inline) =>
                VisitNested(inline.Inlines, _fmt.WithBold().WithItalic());

            protected override void VisitStrikethroughSequenceInline(Omd.StrikethroughSequenceInline inline) =>
                VisitNested(inline.Inlines, _fmt.WithStrike());

            protected override void VisitHighlightSequenceInline(Omd.HighlightSequenceInline inline) =>
                VisitNested(inline.Inlines, _fmt.WithHighlight(HighlightColorValues.Yellow));

            protected override void VisitHtmlTagSequenceInline(Omd.HtmlTagSequenceInline inline) {
                switch (inline.TagName) {
                    case "u":
                    case "ins":
                        VisitNested(inline.Inlines, _fmt.WithUnderline(UnderlineValues.Single));
                        break;
                    case "sup":
                        VisitNested(inline.Inlines, _fmt.WithVerticalTextAlignment(VerticalPositionValues.Superscript));
                        break;
                    case "sub":
                        VisitNested(inline.Inlines, _fmt.WithVerticalTextAlignment(VerticalPositionValues.Subscript));
                        break;
                    case "q":
                        _runs.Add(CreateDetachedRun(_document, "\"", _fmt, _defaultFont));
                        VisitNested(inline.Inlines, _fmt);
                        _runs.Add(CreateDetachedRun(_document, "\"", _fmt, _defaultFont));
                        break;
                    default:
                        VisitNested(inline.Inlines, _fmt);
                        break;
                }
            }

            protected override void VisitHtmlRawInline(Omd.HtmlRawInline inline) {
                if (!string.IsNullOrEmpty(inline.Html)) {
                    _runs.Add(CreateDetachedRun(_document, inline.Html, _fmt, _defaultFont));
                }
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
            new ParagraphInlineWriter(
                paragraph,
                options,
                document,
                footnoteDefs,
                CreateDefaultInlineFormatState(),
                defaultFont)
                .Visit(inlines);
        }

        private sealed class ParagraphInlineWriter : Omd.MarkdownVisitor {
            private readonly WordParagraph _paragraph;
            private readonly MarkdownToWordOptions _options;
            private readonly WordDocument _document;
            private readonly IReadOnlyDictionary<string, string>? _footnoteDefs;
            private readonly InlineFormatState _fmt;
            private readonly string? _defaultFont;

            public ParagraphInlineWriter(
                WordParagraph paragraph,
                MarkdownToWordOptions options,
                WordDocument document,
                IReadOnlyDictionary<string, string>? footnoteDefs,
                InlineFormatState fmt,
                string? defaultFont) {
                _paragraph = paragraph;
                _options = options;
                _document = document;
                _footnoteDefs = footnoteDefs;
                _fmt = fmt;
                _defaultFont = defaultFont;
            }

            private void VisitNested(Omd.MarkdownObject? node, InlineFormatState format) {
                if (node == null) {
                    return;
                }

                new ParagraphInlineWriter(_paragraph, _options, _document, _footnoteDefs, format, _defaultFont).Visit(node);
            }

            protected override void VisitTextRun(Omd.TextRun inline) =>
                AddRun(_paragraph, inline.Text, _fmt, _defaultFont);

            protected override void VisitHardBreakInline(Omd.HardBreakInline inline) =>
                _paragraph.AddBreak();

            protected override void VisitCodeSpanInline(Omd.CodeSpanInline inline) {
                var run = AddRun(_paragraph, inline.Text, _fmt, _defaultFont);
                var mono = FontResolver.Resolve("monospace") ?? "Consolas";
                run.SetFontFamily(mono);
            }

            protected override void VisitLinkInline(Omd.LinkInline inline) {
                try {
                    var uri = new Uri(inline.Url, UriKind.RelativeOrAbsolute);
                    if (inline.LabelInlines != null && inline.LabelInlines.Nodes.Count > 0) {
                        var labelRuns = BuildLinkLabelRunsOmd(inline.LabelInlines, _document, _fmt, _defaultFont);
                        if (labelRuns.Count > 0) {
                            WordHyperLink.AddHyperLink(_paragraph, labelRuns, uri);
                            return;
                        }
                    }

                    var hyperlink = _paragraph.AddHyperLink(inline.Text, uri);
                    ApplyHyperlinkFormattingOmd(hyperlink, _fmt, _defaultFont);
                } catch (UriFormatException ex) {
                    _options.OnWarning?.Invoke($"Invalid URI '{inline.Url}' - emitting as text. {ex.Message}");
                    if (inline.LabelInlines != null && inline.LabelInlines.Nodes.Count > 0) {
                        VisitNested(inline.LabelInlines, _fmt);
                    } else {
                        AddRun(_paragraph, inline.Text, _fmt, _defaultFont);
                    }
                }
            }

            protected override void VisitImageLinkInline(Omd.ImageLinkInline inline) {
                var linkUrl = inline.LinkUrl ?? string.Empty;
                var label = inline.Alt ?? inline.ImageUrl ?? linkUrl;
                try {
                    if (string.IsNullOrEmpty(linkUrl)) {
                        AddRun(_paragraph, label, _fmt, _defaultFont);
                        return;
                    }

                    var uri = new Uri(linkUrl, UriKind.RelativeOrAbsolute);
                    var hyperlink = _paragraph.AddHyperLink(label, uri);
                    ApplyHyperlinkFormattingOmd(hyperlink, _fmt, _defaultFont);
                } catch (UriFormatException ex) {
                    _options.OnWarning?.Invoke($"Invalid URI '{linkUrl}' - emitting alt text. {ex.Message}");
                    AddRun(_paragraph, label, _fmt, _defaultFont);
                }
            }

            protected override void VisitImageInline(Omd.ImageInline inline) =>
                AddRun(_paragraph, inline.Alt ?? string.Empty, _fmt, _defaultFont);

            protected override void VisitFootnoteRefInline(Omd.FootnoteRefInline inline) {
                string text = inline.Label;
                if (_footnoteDefs != null && _footnoteDefs.TryGetValue(inline.Label, out var body)) {
                    text = body;
                }

                _paragraph.AddFootNote(text);
            }

            protected override void VisitBoldInline(Omd.BoldInline inline) =>
                AddRun(_paragraph, inline.Text, _fmt.WithBold(), _defaultFont);

            protected override void VisitItalicInline(Omd.ItalicInline inline) =>
                AddRun(_paragraph, inline.Text, _fmt.WithItalic(), _defaultFont);

            protected override void VisitBoldItalicInline(Omd.BoldItalicInline inline) =>
                AddRun(_paragraph, inline.Text, _fmt.WithBold().WithItalic(), _defaultFont);

            protected override void VisitStrikethroughInline(Omd.StrikethroughInline inline) =>
                AddRun(_paragraph, inline.Text, _fmt.WithStrike(), _defaultFont);

            protected override void VisitHighlightInline(Omd.HighlightInline inline) =>
                AddRun(_paragraph, inline.Text, _fmt.WithHighlight(HighlightColorValues.Yellow), _defaultFont);

            protected override void VisitUnderlineInline(Omd.UnderlineInline inline) =>
                AddRun(_paragraph, inline.Text, _fmt.WithUnderline(UnderlineValues.Single), _defaultFont);

            protected override void VisitBoldSequenceInline(Omd.BoldSequenceInline inline) =>
                VisitNested(inline.Inlines, _fmt.WithBold());

            protected override void VisitItalicSequenceInline(Omd.ItalicSequenceInline inline) =>
                VisitNested(inline.Inlines, _fmt.WithItalic());

            protected override void VisitBoldItalicSequenceInline(Omd.BoldItalicSequenceInline inline) =>
                VisitNested(inline.Inlines, _fmt.WithBold().WithItalic());

            protected override void VisitStrikethroughSequenceInline(Omd.StrikethroughSequenceInline inline) =>
                VisitNested(inline.Inlines, _fmt.WithStrike());

            protected override void VisitHighlightSequenceInline(Omd.HighlightSequenceInline inline) =>
                VisitNested(inline.Inlines, _fmt.WithHighlight(HighlightColorValues.Yellow));

            protected override void VisitHtmlTagSequenceInline(Omd.HtmlTagSequenceInline inline) {
                switch (inline.TagName) {
                    case "u":
                    case "ins":
                        VisitNested(inline.Inlines, _fmt.WithUnderline(UnderlineValues.Single));
                        break;
                    case "sup":
                        VisitNested(inline.Inlines, _fmt.WithVerticalTextAlignment(VerticalPositionValues.Superscript));
                        break;
                    case "sub":
                        VisitNested(inline.Inlines, _fmt.WithVerticalTextAlignment(VerticalPositionValues.Subscript));
                        break;
                    case "q":
                        AddRun(_paragraph, "\"", _fmt, _defaultFont);
                        VisitNested(inline.Inlines, _fmt);
                        AddRun(_paragraph, "\"", _fmt, _defaultFont);
                        break;
                    default:
                        VisitNested(inline.Inlines, _fmt);
                        break;
                }
            }

            protected override void VisitHtmlRawInline(Omd.HtmlRawInline inline) {
                if (!string.IsNullOrEmpty(inline.Html)) {
                    AddRun(_paragraph, inline.Html, _fmt, _defaultFont);
                }
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
