using OfficeIMO.Markdown;
using OfficeIMO.Markup;
using OfficeIMO.Markup.PowerPoint;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Workflows;

internal static partial class MarkdownPowerPointProjection {
    internal static OfficeMarkupDocument Build(MarkdownDocument source, MarkdownToPowerPointOptions options,
        List<OfficeDiagnostic> findings, CancellationToken token) {
        var target = new OfficeMarkupDocument(OfficeMarkupProfile.Presentation);
        var current = new OfficeMarkupSlideBlock();
        bool HasContent() => current.Title != null || current.Blocks.Count > 0 || !string.IsNullOrEmpty(current.Notes);
        void Finish() { if (HasContent()) target.Blocks.Add(current); current = new OfficeMarkupSlideBlock(); }
        if (options.RenderFrontMatter) {
            string? title = source.FindFrontMatterEntry("title")?.Value as string;
            string? subtitle = source.FindFrontMatterEntry("subtitle")?.Value as string;
            string? author = source.FindFrontMatterEntry("author")?.Value as string;
            if (!string.IsNullOrEmpty(title)) { current.Title = title; target.Metadata["title"] = title; }
            if (!string.IsNullOrEmpty(subtitle)) current.Blocks.Add(new OfficeMarkupParagraphBlock(subtitle));
            if (!string.IsNullOrEmpty(author)) { current.Blocks.Add(new OfficeMarkupParagraphBlock(author)); target.Metadata["author"] = author; }
        }
        foreach (IMarkdownBlock block in source.Blocks) {
            token.ThrowIfCancellationRequested();
            if (block is FrontMatterBlock) continue;
            if (block is HorizontalRuleBlock && options.SlideBreak != MarkdownSlideBreak.Heading) { Finish(); continue; }
            if (block is HeadingBlock heading && heading.Level == options.TitleHeadingLevel) {
                if (options.SlideBreak != MarkdownSlideBreak.ThematicBreak && HasContent()) Finish();
                if (current.Title == null) {
                    InspectInlines(heading, null, findings, token, slideTitle: true);
                    current.Title = heading.Text;
                    continue;
                }
            }
            if (block is QuoteBlock quote && options.BlockquotesAsNotes) {
                InspectInlines(quote, null, findings, token, notes: true);
                string notes = quote.ChildBlocks.Count > 0 ? string.Join("\n\n", quote.ChildBlocks.Select(child => child.RenderMarkdown()))
                    : string.Join("\n", quote.Lines);
                current.Notes = string.IsNullOrEmpty(current.Notes) ? notes : current.Notes + "\n\n" + notes;
                continue;
            }
            MarkdownLoadOptions reading = options.ReaderOptions?.Clone() ?? MarkdownLoadOptions.CreateOfficeIMOProfile();
            var diagnostics = new List<OfficeMarkupDiagnostic>();
            MapVisibleBlocks(new[] { block }, current.Blocks, diagnostics, reading, findings, token);
            findings.AddRange(new OfficeMarkupConversionReport(diagnostics).Diagnostics);
        }
        Finish();
        return target;
    }

    private static void MapVisibleBlocks(IEnumerable<IMarkdownBlock> blocks, IList<OfficeMarkupBlock> target,
        IList<OfficeMarkupDiagnostic> diagnostics, MarkdownLoadOptions reading, List<OfficeDiagnostic> findings, CancellationToken token) {
        foreach (IMarkdownBlock block in blocks) {
            token.ThrowIfCancellationRequested();
            if (block is QuoteBlock quote) {
                if (quote.ChildBlocks.Count > 0) MapVisibleBlocks(quote.ChildBlocks, target, diagnostics, reading, findings, token);
                else target.Add(new OfficeMarkupParagraphBlock(string.Join("\n", quote.Lines)));
            } else {
                int start = target.Count;
                OfficeMarkupParser.MapMarkdownBlocks(new[] { block }, target, OfficeMarkupProfile.Presentation, diagnostics, reading);
                if (block is MarkdownObject source && target.Count > start) InspectInlines(source, target[start], findings, token);
            }
        }
    }

    // Markup text blocks have one style, not an inline-run model. Keep common formatting
    // on that style and diagnose semantics that the native plain-text projection discards.
    private static void InspectInlines(MarkdownObject source, OfficeMarkupBlock? target, List<OfficeDiagnostic> findings,
        CancellationToken token, bool notes = false, bool slideTitle = false) {
        var inspection = new InlineProjectionVisitor(findings, token, notes, source.GetType().Name);
        inspection.Visit(source);
        IDictionary<string, string>? attributes = target switch {
            OfficeMarkupParagraphBlock or OfficeMarkupHeadingBlock => target.Attributes,
            OfficeMarkupListBlock list when OfficeMarkupListTraversal.IsTextOnly(list) => list.Attributes,
            _ => null
        };
        InlineStyle preserved = attributes != null ? inspection.CommonStyle
            : slideTitle ? inspection.CommonStyle & InlineStyle.Bold : InlineStyle.None;
        if (attributes != null) {
            if (preserved.HasFlag(InlineStyle.Bold)) attributes["bold"] = "true";
            if (preserved.HasFlag(InlineStyle.Italic)) attributes["italic"] = "true";
            if (preserved.HasFlag(InlineStyle.Underline)) attributes["underline"] = "true";
            if (preserved.HasFlag(InlineStyle.Strikethrough)) attributes["strikethrough"] = "true";
            if (preserved.HasFlag(InlineStyle.Highlight)) attributes["highlight"] = "#FFFF00";
            if (preserved.HasFlag(InlineStyle.Superscript)) attributes["baseline"] = "superscript";
            if (preserved.HasFlag(InlineStyle.Subscript)) attributes["baseline"] = "subscript";
        }
        inspection.ReportStyleLoss(preserved);
    }

    [Flags]
    private enum InlineStyle { None = 0, Bold = 1, Italic = 2, Underline = 4, Strikethrough = 8, Highlight = 16, Superscript = 32, Subscript = 64 }

    private sealed class InlineProjectionVisitor(List<OfficeDiagnostic> findings, CancellationToken token, bool notes, string context) : MarkdownVisitor {
        private readonly List<(MarkdownInline Node, InlineStyle Style)> _styles = new();
        private InlineStyle _active;
        private bool _hasText;
        internal InlineStyle CommonStyle { get; private set; }

        protected override void VisitInline(MarkdownInline inline) {
            token.ThrowIfCancellationRequested();
            InlineStyle previous = _active;
            InlineStyle own = StyleOf(inline);
            if ((own & (InlineStyle.Superscript | InlineStyle.Subscript)) != 0)
                _active &= ~(InlineStyle.Superscript | InlineStyle.Subscript);
            _active |= own;
            if (own != InlineStyle.None) _styles.Add((inline, own));
            switch (inline) {
                case LinkInline link: ReportLink(link.Url); break;
                case ImageLinkInline image: ReportLink(image.LinkUrl); break;
                case CodeSpanInline: ReportSimplification(inline, "Inline code typography"); break;
                case HtmlRawInline: ReportSimplification(inline, "Raw inline HTML presentation"); break;
                case HtmlTagSequenceInline tag when own == InlineStyle.None:
                    ReportSimplification(inline, $"Inline HTML '{tag.TagName}' presentation"); break;
            }
            if (inline.ChildObjects.Count == 0 && inline is IPlainTextMarkdownInline text) {
                var plain = new System.Text.StringBuilder();
                text.AppendPlainText(plain);
                if (!string.IsNullOrWhiteSpace(plain.ToString())) {
                    CommonStyle = _hasText ? CommonStyle & _active : _active;
                    _hasText = true;
                }
            }
            base.VisitInline(inline);
            _active = previous;
        }

        internal void ReportStyleLoss(InlineStyle preserved) {
            foreach (var (node, style) in _styles) {
                InlineStyle lost = style & ~preserved;
                if (lost != InlineStyle.None) ReportSimplification(node, $"Inline {lost} formatting");
            }
        }

        private void ReportLink(string target) => findings.Add(new OfficeDiagnostic(
            notes ? OfficeDiagnosticCodes.SimplifiedBlock : OfficeDiagnosticCodes.ShapeHyperlinkOmitted,
            OfficeDiagnosticSeverity.Warning,
            notes ? $"Hyperlink '{target}' is retained as Markdown source in speaker notes, without an interactive link."
                : $"The hyperlink target '{target}' is omitted from {context}; its visible label is retained.",
            target, "Markdown", notes ? OfficeConversionLossKind.Approximation : OfficeConversionLossKind.Omission));

        private void ReportSimplification(MarkdownInline node, string feature) => findings.Add(new OfficeDiagnostic(
            OfficeDiagnosticCodes.SimplifiedBlock, OfficeDiagnosticSeverity.Warning,
            notes ? $"{feature} is retained as Markdown source in speaker notes instead of rendered inline formatting."
                : $"{feature} in {context} is flattened by the presentation's single-style text projection.",
            node.SourceSpan is { } span ? $"{context}:line {span.StartLine}" : context, "Markdown", OfficeConversionLossKind.Approximation));

        private static InlineStyle StyleOf(MarkdownInline inline) => inline switch {
            BoldInline or BoldSequenceInline => InlineStyle.Bold,
            ItalicInline or ItalicSequenceInline => InlineStyle.Italic,
            BoldItalicInline or BoldItalicSequenceInline => InlineStyle.Bold | InlineStyle.Italic,
            UnderlineInline or InsertedInline or InsertedSequenceInline => InlineStyle.Underline,
            StrikethroughInline or StrikethroughSequenceInline => InlineStyle.Strikethrough,
            HighlightInline or HighlightSequenceInline => InlineStyle.Highlight,
            SuperscriptInline or SuperscriptSequenceInline => InlineStyle.Superscript,
            SubscriptInline or SubscriptSequenceInline => InlineStyle.Subscript,
            HtmlTagSequenceInline tag => tag.TagName switch {
                "strong" or "b" => InlineStyle.Bold,
                "em" or "i" => InlineStyle.Italic,
                "u" or "ins" => InlineStyle.Underline,
                "s" or "del" or "strike" => InlineStyle.Strikethrough,
                "mark" => InlineStyle.Highlight,
                "sup" => InlineStyle.Superscript,
                "sub" => InlineStyle.Subscript,
                _ => InlineStyle.None
            },
            _ => InlineStyle.None
        };
    }

    internal static MarkupPowerPointExportOptions ExportOptions(MarkdownToPowerPointOptions options, string root) {
        var output = new MarkupPowerPointExportOptions { BaseDirectory = root, AllowExternalImagePaths = false,
            AllowDataUriImages = options.AllowDataUriImages, MaximumDataUriImageBytes = options.MaximumDataUriImageBytes,
            RenderMermaidDiagrams = options.RenderMermaidDiagrams, TemporaryDirectory = Path.Combine(root, "mermaid") };
        if (options.SlideSize is { } size) {
            (output.SlideWidthInches, output.SlideHeightInches) = size switch {
                OfficeIMO.PowerPoint.PowerPointSlideSizePreset.Screen4x3 => (10D, 7.5D),
                OfficeIMO.PowerPoint.PowerPointSlideSizePreset.Screen16x9 => (13.333D, 7.5D),
                OfficeIMO.PowerPoint.PowerPointSlideSizePreset.Screen16x10 => (10D, 6.25D),
                _ => throw new ArgumentOutOfRangeException(nameof(options.SlideSize))
            };
        }
        return output;
    }
}
