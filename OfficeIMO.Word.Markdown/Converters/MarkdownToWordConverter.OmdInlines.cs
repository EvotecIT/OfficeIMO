using System.Text;
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
                VerticalPositionValues? verticalTextAlignment,
                StyleRunProperties? nativeDefaults = null) {
                Bold = bold;
                Italic = italic;
                Strike = strike;
                Underline = underline;
                Highlight = highlight;
                VerticalTextAlignment = verticalTextAlignment;
                NativeDefaults = nativeDefaults;
            }

            public bool Bold { get; }
            public bool Italic { get; }
            public bool Strike { get; }
            public UnderlineValues? Underline { get; }
            public HighlightColorValues? Highlight { get; }
            public VerticalPositionValues? VerticalTextAlignment { get; }
            public StyleRunProperties? NativeDefaults { get; }

            public InlineFormatState WithBold() => new InlineFormatState(bold: true, italic: Italic, strike: Strike, underline: Underline, highlight: Highlight, verticalTextAlignment: VerticalTextAlignment, nativeDefaults: NativeDefaults);
            public InlineFormatState WithItalic() => new InlineFormatState(bold: Bold, italic: true, strike: Strike, underline: Underline, highlight: Highlight, verticalTextAlignment: VerticalTextAlignment, nativeDefaults: NativeDefaults);
            public InlineFormatState WithStrike() => new InlineFormatState(bold: Bold, italic: Italic, strike: true, underline: Underline, highlight: Highlight, verticalTextAlignment: VerticalTextAlignment, nativeDefaults: NativeDefaults);
            public InlineFormatState WithUnderline(UnderlineValues underline) => new InlineFormatState(bold: Bold, italic: Italic, strike: Strike, underline: underline, highlight: Highlight, verticalTextAlignment: VerticalTextAlignment, nativeDefaults: NativeDefaults);
            public InlineFormatState WithHighlight(HighlightColorValues highlight) => new InlineFormatState(bold: Bold, italic: Italic, strike: Strike, underline: Underline, highlight: highlight, verticalTextAlignment: VerticalTextAlignment, nativeDefaults: NativeDefaults);
            public InlineFormatState WithVerticalTextAlignment(VerticalPositionValues verticalTextAlignment) => new InlineFormatState(bold: Bold, italic: Italic, strike: Strike, underline: Underline, highlight: Highlight, verticalTextAlignment: verticalTextAlignment, nativeDefaults: NativeDefaults);
        }

        private static WordParagraph AddRun(WordParagraph paragraph, string? text, InlineFormatState fmt, string? defaultFont, string? defaultTextColorHex) {
            var run = paragraph.AddText(text ?? string.Empty);
            ApplyInlineRunFormatting(run, fmt, defaultFont, defaultTextColorHex);
            return run;
        }

        private static void ApplyInlineRunFormatting(WordParagraph run, InlineFormatState fmt, string? defaultFont, string? defaultTextColorHex) {
            ApplyNativeTableRunDefaults(run, fmt, omitFont: !string.IsNullOrEmpty(defaultFont));
            if (fmt.Bold) run.SetBold();
            if (fmt.Italic) run.SetItalic();
            if (fmt.Underline.HasValue && fmt.Underline.Value != UnderlineValues.None) run.SetUnderline(fmt.Underline.Value.ToOfficeEnum());
            if (fmt.Strike) run.SetStrikethrough();
            if (fmt.Highlight.HasValue && fmt.Highlight.Value != HighlightColorValues.None) run.SetHighlight(fmt.Highlight.Value.ToOfficeEnum());
            if (fmt.VerticalTextAlignment.HasValue) run.SetVerticalTextAlignment(fmt.VerticalTextAlignment.Value.ToOfficeEnum());
            if (!string.IsNullOrEmpty(defaultFont)) run.SetFontFamily(defaultFont!);
            if (!string.IsNullOrEmpty(defaultTextColorHex)) run.SetColor(OfficeIMO.OfficeColor.Parse(defaultTextColorHex!));
        }

        // Monospace faces have a larger x-height than proportional body fonts, so code at
        // the body size reads a size bigger. Body paragraphs set code one tenth smaller;
        // headings keep their style size.
        private const double InlineCodeSizeRatio = 0.9D;
        private const int DefaultBodyFontSize = 11;

        private static void ApplyInlineCodeSize(WordParagraph run, WordParagraph paragraph, WordDocument document, StyleRunProperties? nativeDefaults = null) {
            if (run.FontSize.HasValue && nativeDefaults == null) return;
            WordParagraphStyles? style = paragraph.Style;
            if (style.HasValue && style.Value != WordParagraphStyles.Normal && style.Value != WordParagraphStyles.ListParagraph) return;
            double bodySize = nativeDefaults != null && run.FontSize.HasValue
                ? run.FontSize.Value.ToPoints() : document.Settings.FontSize?.ToPoints() ?? DefaultBodyFontSize;
            run.FontSize = OfficeLength.Points(Math.Max(1, (int)Math.Round(bodySize * InlineCodeSizeRatio, MidpointRounding.AwayFromZero)));
        }

        private static WordParagraph CreateDetachedRun(WordDocument document, string? text, InlineFormatState fmt, string? defaultFont, bool forceMonospace = false) {
            var paragraph = new Paragraph();
            var run = new Run(new Text(text ?? string.Empty) {
                Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve
            });
            paragraph.Append(run);

            var wrapper = new WordParagraph(document, paragraph, run);
            ApplyNativeTableRunDefaults(wrapper, fmt, omitFont: forceMonospace || !string.IsNullOrEmpty(defaultFont));
            if (fmt.Bold) wrapper.SetBold();
            if (fmt.Italic) wrapper.SetItalic();
            if (fmt.Underline.HasValue && fmt.Underline.Value != UnderlineValues.None) wrapper.SetUnderline(fmt.Underline.Value.ToOfficeEnum());
            if (fmt.Strike) wrapper.SetStrikethrough();
            if (fmt.Highlight.HasValue && fmt.Highlight.Value != HighlightColorValues.None) wrapper.SetHighlight(fmt.Highlight.Value.ToOfficeEnum());
            if (fmt.VerticalTextAlignment.HasValue) wrapper.SetVerticalTextAlignment(fmt.VerticalTextAlignment.Value.ToOfficeEnum());

            if (forceMonospace) {
                wrapper.SetFontFamily(WordFontResolver.Resolve("monospace") ?? "Consolas");
            } else if (!string.IsNullOrEmpty(defaultFont)) {
                wrapper.SetFontFamily(defaultFont!);
            }

            return wrapper;
        }

        private static List<WordParagraph> BuildLinkLabelRunsOmd(
            Omd.InlineSequence? inlines,
            WordDocument document,
            InlineFormatState fmt,
            string? defaultFont,
            int maxNestingDepth,
            int nestingDepth) {
            var runs = new List<WordParagraph>();
            if (inlines == null) {
                return runs;
            }

            EnsureInlineNestingWithinLimit(maxNestingDepth, nestingDepth);

            new DetachedInlineRunBuilder(
                document, runs, fmt, defaultFont, maxNestingDepth,
                nestingDepth + 1)
                .Visit(inlines);
            return runs;
        }

        private static bool TryGetPlainLeafInlineText(Omd.MarkdownInline inline, out string text) {
            if (inline is Omd.IPlainTextMarkdownInline plain && inline.ChildObjects.Count == 0) {
                var builder = new StringBuilder();
                plain.AppendPlainText(builder);
                text = builder.ToString();
                return true;
            }

            text = string.Empty;
            return false;
        }

        private sealed class DetachedInlineRunBuilder : Omd.MarkdownVisitor {
            private readonly WordDocument _document;
            private readonly List<WordParagraph> _runs;
            private readonly InlineFormatState _fmt;
            private readonly string? _defaultFont;
            private readonly int _maxNestingDepth;
            private readonly int _nestingDepth;

            public DetachedInlineRunBuilder(
                WordDocument document,
                List<WordParagraph> runs,
                InlineFormatState fmt,
                string? defaultFont,
                int maxNestingDepth,
                int nestingDepth) {
                _document = document;
                _runs = runs;
                _fmt = fmt;
                _defaultFont = defaultFont;
                _maxNestingDepth = maxNestingDepth;
                _nestingDepth = nestingDepth;
            }

            private void VisitNested(Omd.MarkdownObject? node, InlineFormatState format) {
                if (node == null) {
                    return;
                }
                EnsureInlineNestingWithinLimit(
                    _maxNestingDepth, _nestingDepth);

                new DetachedInlineRunBuilder(
                    _document,
                    _runs,
                    format,
                    _defaultFont,
                    _maxNestingDepth,
                    _nestingDepth + 1)
                    .Visit(node);
            }

            protected override void VisitTextRun(Omd.MarkdownTextRun inline) =>
                _runs.Add(CreateDetachedRun(_document, inline.Text, _fmt, _defaultFont));

            protected override void VisitHardBreakInline(Omd.HardBreakInline inline) =>
                _runs.Add(CreateDetachedRun(_document, " ", _fmt, _defaultFont));

            protected override void VisitSoftBreakInline(Omd.SoftBreakInline inline) =>
                _runs.Add(CreateDetachedRun(_document, " ", _fmt, _defaultFont));

            protected override void VisitCodeSpanInline(Omd.CodeSpanInline inline) =>
                _runs.Add(CreateDetachedRun(_document, inline.Text, _fmt, _defaultFont, forceMonospace: true));

            protected override void VisitLinkInline(Omd.LinkInline inline) =>
                _runs.Add(CreateDetachedRun(_document, inline.Text, _fmt, _defaultFont));

            protected override void VisitImageInline(Omd.ImageInline inline) =>
                _runs.Add(CreateDetachedRun(_document, inline.Alt ?? string.Empty, _fmt, _defaultFont));

            protected override void VisitAbbreviationInline(Omd.AbbreviationInline inline) =>
                _runs.Add(CreateDetachedRun(_document, inline.Text, _fmt, _defaultFont));

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

            protected override void VisitInsertedInline(Omd.InsertedInline inline) =>
                _runs.Add(CreateDetachedRun(_document, inline.Text, _fmt.WithUnderline(UnderlineValues.Single), _defaultFont));

            protected override void VisitSuperscriptInline(Omd.SuperscriptInline inline) =>
                _runs.Add(CreateDetachedRun(_document, inline.Text, _fmt.WithVerticalTextAlignment(VerticalPositionValues.Superscript), _defaultFont));

            protected override void VisitSubscriptInline(Omd.SubscriptInline inline) =>
                _runs.Add(CreateDetachedRun(_document, inline.Text, _fmt.WithVerticalTextAlignment(VerticalPositionValues.Subscript), _defaultFont));

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

            protected override void VisitInsertedSequenceInline(Omd.InsertedSequenceInline inline) =>
                VisitNested(inline.Inlines, _fmt.WithUnderline(UnderlineValues.Single));

            protected override void VisitSuperscriptSequenceInline(Omd.SuperscriptSequenceInline inline) =>
                VisitNested(inline.Inlines, _fmt.WithVerticalTextAlignment(VerticalPositionValues.Superscript));

            protected override void VisitSubscriptSequenceInline(Omd.SubscriptSequenceInline inline) =>
                VisitNested(inline.Inlines, _fmt.WithVerticalTextAlignment(VerticalPositionValues.Subscript));

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

            protected override void VisitInline(Omd.MarkdownInline inline) {
                if (TryGetPlainLeafInlineText(inline, out var text)) {
                    _runs.Add(CreateDetachedRun(_document, text, _fmt, _defaultFont));
                    return;
                }

                base.VisitInline(inline);
            }
        }

        private static InlineFormatState CreateDefaultInlineFormatState(StyleRunProperties? nativeDefaults = null) =>
            new InlineFormatState(
                bold: false,
                italic: false,
                strike: false,
                underline: null,
                highlight: null,
                verticalTextAlignment: null,
                nativeDefaults: nativeDefaults);

        /// <summary>
        /// Processes OfficeIMO.Markdown inline sequence into Word runs.
        /// </summary>
        private static void ProcessInlinesOmd(
            Omd.InlineSequence inlines,
            WordParagraph paragraph,
            MarkdownWordExportOptions options,
            WordDocument document,
            IReadOnlyDictionary<string, string>? footnoteDefs = null,
            double pageContentWidthPixels = 0,
            int listLevel = 0,
            int quoteDepth = 0
        ) {
            if (inlines == null) return;

            string? defaultFont = ResolveDefaultFontFamily(options);
            Omd.MarkdownVisualTheme? theme = ResolveTheme(options);
            StyleRunProperties? nativeDefaults = PrepareNativeTableTextDefaults(paragraph, theme);
            string? defaultTextColorHex = theme?.SourceTheme == null ? theme?.PaletteSnapshot.Text.ToRgbHex() : null;
            new ParagraphInlineWriter(
                paragraph,
                options,
                document,
                footnoteDefs,
                CreateDefaultInlineFormatState(nativeDefaults),
                defaultFont,
                defaultTextColorHex,
                pageContentWidthPixels,
                listLevel,
                quoteDepth,
                nestingDepth: 0)
                .Visit(inlines);
        }

        private sealed class ParagraphInlineWriter : Omd.MarkdownVisitor {
            private readonly WordParagraph _paragraph;
            private readonly MarkdownWordExportOptions _options;
            private readonly WordDocument _document;
            private readonly IReadOnlyDictionary<string, string>? _footnoteDefs;
            private readonly InlineFormatState _fmt;
            private readonly string? _defaultFont;
            private readonly string? _defaultTextColorHex;
            private readonly double _pageContentWidthPixels;
            private readonly int _listLevel;
            private readonly int _quoteDepth;
            private readonly int _nestingDepth;

            public ParagraphInlineWriter(
                WordParagraph paragraph,
                MarkdownWordExportOptions options,
                WordDocument document,
                IReadOnlyDictionary<string, string>? footnoteDefs,
                InlineFormatState fmt,
                string? defaultFont,
                string? defaultTextColorHex,
                double pageContentWidthPixels,
                int listLevel,
                int quoteDepth,
                int nestingDepth) {
                _paragraph = paragraph;
                _options = options;
                _document = document;
                _footnoteDefs = footnoteDefs;
                _fmt = fmt;
                _defaultFont = defaultFont;
                _defaultTextColorHex = defaultTextColorHex;
                _pageContentWidthPixels = pageContentWidthPixels;
                _listLevel = listLevel;
                _quoteDepth = quoteDepth;
                _nestingDepth = nestingDepth;
            }

            private void VisitNested(Omd.MarkdownObject? node, InlineFormatState format) {
                if (node == null) {
                    return;
                }
                EnsureInlineNestingWithinLimit(
                    _options.MaxNestingDepth, _nestingDepth);

                new ParagraphInlineWriter(
                    _paragraph,
                    _options,
                    _document,
                    _footnoteDefs,
                    format,
                    _defaultFont,
                    _defaultTextColorHex,
                    _pageContentWidthPixels,
                    _listLevel,
                    _quoteDepth,
                    _nestingDepth + 1)
                    .Visit(node);
            }

            protected override void VisitTextRun(Omd.MarkdownTextRun inline) =>
                AddRun(_paragraph, inline.Text, _fmt, _defaultFont, _defaultTextColorHex);

            protected override void VisitHardBreakInline(Omd.HardBreakInline inline) =>
                ApplyNativeTableRunDefaults(_paragraph.AddBreak(), _fmt, omitFont: !string.IsNullOrEmpty(_defaultFont));

            protected override void VisitSoftBreakInline(Omd.SoftBreakInline inline) =>
                AddRun(_paragraph, " ", _fmt, _defaultFont, _defaultTextColorHex);

            protected override void VisitCodeSpanInline(Omd.CodeSpanInline inline) {
                var run = AddRun(_paragraph, inline.Text, _fmt, _defaultFont, _defaultTextColorHex);
                var mono = WordFontResolver.Resolve("monospace") ?? "Consolas";
                if (run.ScopedRunProperties is { } properties) properties.RunFonts = null;
                run.SetFontFamily(mono);
                ApplyInlineCodeSize(run, _paragraph, _document, _fmt.NativeDefaults);
            }

            protected override void VisitLinkInline(Omd.LinkInline inline) {
                string? internalAnchor = ResolveInternalLinkBookmark(inline.Url, _options);
                if (internalAnchor != null) {
                    if (inline.LabelInlines != null && inline.LabelInlines.Nodes.Count > 0) {
                        var anchorLabelRuns = BuildLinkLabelRunsOmd(
                            inline.LabelInlines,
                            _document,
                            _fmt,
                            _defaultFont,
                            _options.MaxNestingDepth,
                            _nestingDepth);
                        if (anchorLabelRuns.Count > 0) {
                            ApplyHyperlinkRunFormattingOmd(anchorLabelRuns, _options);
                            WordHyperLink.AddHyperLink(_paragraph, anchorLabelRuns, internalAnchor);
                            return;
                        }
                    }

                    var anchorLink = _paragraph.AddHyperLink(inline.Text, internalAnchor);
                    ApplyHyperlinkFormattingOmd(anchorLink.Paragraph, _fmt, _defaultFont, _options);
                    return;
                }

                try {
                    var uri = new Uri(inline.Url, UriKind.RelativeOrAbsolute);
                    if (inline.LabelInlines != null && inline.LabelInlines.Nodes.Count > 0) {
                        var labelRuns = BuildLinkLabelRunsOmd(
                            inline.LabelInlines,
                            _document,
                            _fmt,
                            _defaultFont,
                            _options.MaxNestingDepth,
                            _nestingDepth);
                        if (labelRuns.Count > 0) {
                            ApplyHyperlinkRunFormattingOmd(labelRuns, _options);
                            WordHyperLink.AddHyperLink(_paragraph, labelRuns, uri);
                            return;
                        }
                    }

                    var hyperlink = _paragraph.AddHyperLink(inline.Text, uri).Paragraph;
                    ApplyHyperlinkFormattingOmd(hyperlink, _fmt, _defaultFont, _options);
                } catch (UriFormatException ex) {
                    _options.OnWarning?.Invoke($"Invalid URI '{inline.Url}' - emitting as text. {ex.Message}");
                    if (inline.LabelInlines != null && inline.LabelInlines.Nodes.Count > 0) {
                        VisitNested(inline.LabelInlines, _fmt);
                    } else {
                        AddRun(_paragraph, inline.Text, _fmt, _defaultFont, _defaultTextColorHex);
                    }
                }
            }

            protected override void VisitImageLinkInline(Omd.ImageLinkInline inline) {
                var linkUrl = inline.LinkUrl ?? string.Empty;
                var label = inline.Alt ?? inline.ImageUrl ?? linkUrl;
                try {
                    if (string.IsNullOrEmpty(linkUrl)) {
                        AddRun(_paragraph, label, _fmt, _defaultFont, _defaultTextColorHex);
                        return;
                    }

                    var uri = new Uri(linkUrl, UriKind.RelativeOrAbsolute);
                    var hyperlink = _paragraph.AddHyperLink(label, uri).Paragraph;
                    ApplyHyperlinkFormattingOmd(hyperlink, _fmt, _defaultFont, _options);
                } catch (UriFormatException ex) {
                    _options.OnWarning?.Invoke($"Invalid URI '{linkUrl}' - emitting alt text. {ex.Message}");
                    AddRun(_paragraph, label, _fmt, _defaultFont, _defaultTextColorHex);
                }
            }

            protected override void VisitImageInline(Omd.ImageInline inline) =>
                RenderMarkdownImageIntoParagraph(
                    _paragraph,
                    inline.Src,
                    inline.PlainAlt,
                    null,
                    null,
                    _options,
                    _pageContentWidthPixels,
                    _listLevel,
                    _quoteDepth,
                    "inline",
                    _fmt);

            protected override void VisitAbbreviationInline(Omd.AbbreviationInline inline) =>
                AddRun(_paragraph, inline.Text, _fmt, _defaultFont, _defaultTextColorHex);

            protected override void VisitFootnoteRefInline(Omd.FootnoteRefInline inline) {
                string text = inline.Label;
                if (_footnoteDefs != null && _footnoteDefs.TryGetValue(inline.Label, out var body)) {
                    text = body;
                }

                var reference = _paragraph.AddFootNote(text).Paragraph;
                ApplyNativeContextRunFormatting(reference, _paragraph, _options, _fmt);
            }

            protected override void VisitBoldInline(Omd.BoldInline inline) =>
                AddRun(_paragraph, inline.Text, _fmt.WithBold(), _defaultFont, _defaultTextColorHex);

            protected override void VisitItalicInline(Omd.ItalicInline inline) =>
                AddRun(_paragraph, inline.Text, _fmt.WithItalic(), _defaultFont, _defaultTextColorHex);

            protected override void VisitBoldItalicInline(Omd.BoldItalicInline inline) =>
                AddRun(_paragraph, inline.Text, _fmt.WithBold().WithItalic(), _defaultFont, _defaultTextColorHex);

            protected override void VisitStrikethroughInline(Omd.StrikethroughInline inline) =>
                AddRun(_paragraph, inline.Text, _fmt.WithStrike(), _defaultFont, _defaultTextColorHex);

            protected override void VisitHighlightInline(Omd.HighlightInline inline) =>
                AddRun(_paragraph, inline.Text, _fmt.WithHighlight(HighlightColorValues.Yellow), _defaultFont, _defaultTextColorHex);

            protected override void VisitUnderlineInline(Omd.UnderlineInline inline) =>
                AddRun(_paragraph, inline.Text, _fmt.WithUnderline(UnderlineValues.Single), _defaultFont, _defaultTextColorHex);

            protected override void VisitInsertedInline(Omd.InsertedInline inline) =>
                AddRun(_paragraph, inline.Text, _fmt.WithUnderline(UnderlineValues.Single), _defaultFont, _defaultTextColorHex);

            protected override void VisitSuperscriptInline(Omd.SuperscriptInline inline) =>
                AddRun(_paragraph, inline.Text, _fmt.WithVerticalTextAlignment(VerticalPositionValues.Superscript), _defaultFont, _defaultTextColorHex);

            protected override void VisitSubscriptInline(Omd.SubscriptInline inline) =>
                AddRun(_paragraph, inline.Text, _fmt.WithVerticalTextAlignment(VerticalPositionValues.Subscript), _defaultFont, _defaultTextColorHex);

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

            protected override void VisitInsertedSequenceInline(Omd.InsertedSequenceInline inline) =>
                VisitNested(inline.Inlines, _fmt.WithUnderline(UnderlineValues.Single));

            protected override void VisitSuperscriptSequenceInline(Omd.SuperscriptSequenceInline inline) =>
                VisitNested(inline.Inlines, _fmt.WithVerticalTextAlignment(VerticalPositionValues.Superscript));

            protected override void VisitSubscriptSequenceInline(Omd.SubscriptSequenceInline inline) =>
                VisitNested(inline.Inlines, _fmt.WithVerticalTextAlignment(VerticalPositionValues.Subscript));

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
                        AddRun(_paragraph, "\"", _fmt, _defaultFont, _defaultTextColorHex);
                        VisitNested(inline.Inlines, _fmt);
                        AddRun(_paragraph, "\"", _fmt, _defaultFont, _defaultTextColorHex);
                        break;
                    default:
                        VisitNested(inline.Inlines, _fmt);
                        break;
                }
            }

            protected override void VisitHtmlRawInline(Omd.HtmlRawInline inline) {
                if (!string.IsNullOrEmpty(inline.Html)) {
                    AddRun(_paragraph, inline.Html, _fmt, _defaultFont, _defaultTextColorHex);
                }
            }

            protected override void VisitInline(Omd.MarkdownInline inline) {
                if (TryGetPlainLeafInlineText(inline, out var text)) {
                    AddRun(_paragraph, text, _fmt, _defaultFont, _defaultTextColorHex);
                    return;
                }

                base.VisitInline(inline);
            }
        }

        private static void EnsureInlineNestingWithinLimit(
            int maxNestingDepth,
            int currentDepth) {
            if (maxNestingDepth < 1) {
                throw new ArgumentOutOfRangeException(
                    nameof(MarkdownWordExportOptions.MaxNestingDepth),
                    maxNestingDepth,
                    "MaxNestingDepth must be greater than zero.");
            }
            if (currentDepth >= maxNestingDepth) {
                throw new System.IO.InvalidDataException(
                    $"Markdown-to-Word inline nesting exceeds MaxNestingDepth ({maxNestingDepth}).");
            }
        }

        private static void ApplyHyperlinkFormattingOmd(WordParagraph hyperlink, InlineFormatState fmt, string? defaultFont, MarkdownWordExportOptions options) {
            ApplyNativeTableRunDefaults(hyperlink, fmt, omitFont: !string.IsNullOrEmpty(defaultFont));
            if (fmt.Bold) hyperlink.SetBold();
            if (fmt.Italic) hyperlink.SetItalic();
            if (fmt.Underline.HasValue && fmt.Underline.Value != UnderlineValues.None) hyperlink.SetUnderline(fmt.Underline.Value.ToOfficeEnum());
            if (fmt.Strike) hyperlink.SetStrikethrough();
            if (fmt.Highlight.HasValue && fmt.Highlight.Value != HighlightColorValues.None) hyperlink.SetHighlight(fmt.Highlight.Value.ToOfficeEnum());
            if (fmt.VerticalTextAlignment.HasValue) hyperlink.SetVerticalTextAlignment(fmt.VerticalTextAlignment.Value.ToOfficeEnum());
            if (!string.IsNullOrEmpty(defaultFont)) hyperlink.SetFontFamily(defaultFont!);
            string? accentHex = ResolveLinkColor(options);
            if (!string.IsNullOrEmpty(accentHex)) hyperlink.SetColor(OfficeIMO.OfficeColor.Parse(accentHex!));
        }

        private static void ApplyHyperlinkRunFormattingOmd(IEnumerable<WordParagraph> runs, MarkdownWordExportOptions options) {
            string? accentHex = ResolveLinkColor(options);
            if (string.IsNullOrEmpty(accentHex)) {
                return;
            }

            foreach (WordParagraph run in runs) {
                run.SetColor(OfficeIMO.OfficeColor.Parse(accentHex!));
            }
        }

        private static string? ResolveLinkColor(MarkdownWordExportOptions options) {
            Omd.MarkdownVisualTheme? theme = ResolveTheme(options);
            return theme?.EffectiveSourceTheme is { } source
                ? source.TextStyle.LinkColor.Resolve(source).ToRgbHex()
                : theme?.PaletteSnapshot.Accent.ToRgbHex();
        }
    }
}
