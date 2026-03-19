namespace OfficeIMO.Markdown;

internal static class MarkdownInlineSyntaxBuilder {
    internal static IReadOnlyList<MarkdownSyntaxNode> BuildChildren(InlineSequence? sequence) {
        if (sequence == null || sequence.Nodes.Count == 0) {
            return Array.Empty<MarkdownSyntaxNode>();
        }

        var nodes = new List<MarkdownSyntaxNode>(sequence.Nodes.Count);
        for (int i = 0; i < sequence.Nodes.Count; i++) {
            var node = BuildNode(sequence.Nodes[i]);
            if (node != null) {
                nodes.Add(node);
            }
        }

        return nodes;
    }

    private static MarkdownSyntaxNode? BuildNode(IMarkdownInline? inline) {
        if (inline == null) {
            return null;
        }

        var span = MarkdownInlineSourceSpans.Get(inline);
        switch (inline) {
            case TextRun text:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineText, span, literal: text.Text);
            case DecodedHtmlEntityTextRun text:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineText, span, literal: text.Text);
            case CodeSpanInline code:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineCodeSpan, span, literal: code.Text);
            case FootnoteRefInline footnote:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineFootnoteRef, span, literal: footnote.Label);
            case HardBreakInline:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineHardBreak, span, literal: "\\n");
            case LinkInline link:
                return new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.InlineLink,
                    span,
                    literal: link.Url,
                    children: BuildInlineLabelChildren(link.LabelInlines, link.Text));
            case ImageLinkInline imageLink:
                return new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.InlineImageLink,
                    span,
                    literal: imageLink.LinkUrl,
                    children: BuildInlineImageChildren(imageLink.Alt, imageLink.ImageUrl, imageLink.Title, imageLink.LinkTitle));
            case ImageInline image:
                return new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.InlineImage,
                    span,
                    literal: image.Src,
                    children: BuildInlineImageChildren(image.Alt, image.Src, image.Title, null));
            case BoldSequenceInline bold:
                return BuildContainerNode(MarkdownSyntaxKind.InlineStrong, bold.Inlines);
            case BoldInline bold:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineStrong, span, literal: bold.Text);
            case ItalicSequenceInline italic:
                return BuildContainerNode(MarkdownSyntaxKind.InlineEmphasis, italic.Inlines);
            case ItalicInline italic:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineEmphasis, span, literal: italic.Text);
            case BoldItalicSequenceInline boldItalic:
                return BuildContainerNode(MarkdownSyntaxKind.InlineStrongEmphasis, boldItalic.Inlines);
            case BoldItalicInline boldItalic:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineStrongEmphasis, span, literal: boldItalic.Text);
            case StrikethroughSequenceInline strike:
                return BuildContainerNode(MarkdownSyntaxKind.InlineStrikethrough, strike.Inlines);
            case StrikethroughInline strike:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineStrikethrough, span, literal: strike.Text);
            case HighlightSequenceInline highlight:
                return BuildContainerNode(MarkdownSyntaxKind.InlineHighlight, highlight.Inlines);
            case HighlightInline highlight:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineHighlight, span, literal: highlight.Text);
            case UnderlineInline underline:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineUnderline, span, literal: underline.Text);
            case HtmlTagSequenceInline htmlTag:
                return new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.InlineHtmlTag,
                    span,
                    literal: htmlTag.TagName,
                    children: BuildChildren(htmlTag.Inlines));
            case HtmlRawInline html:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineHtmlRaw, span, literal: html.Html);
            case IInlineContainerMarkdownInline container when container.NestedInlines != null:
                return new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.Unknown,
                    span ?? MarkdownBlockSyntaxBuilder.GetAggregateSpan(BuildChildren(container.NestedInlines)),
                    literal: ((IRenderableMarkdownInline)inline).RenderMarkdown(),
                    children: BuildChildren(container.NestedInlines));
            case IRenderableMarkdownInline renderable:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.Unknown, span, literal: renderable.RenderMarkdown());
            default:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.Unknown, span, literal: inline.ToString());
        }
    }

    private static MarkdownSyntaxNode BuildContainerNode(MarkdownSyntaxKind kind, InlineSequence sequence) {
        var children = BuildChildren(sequence);
        return new MarkdownSyntaxNode(kind, MarkdownBlockSyntaxBuilder.GetAggregateSpan(children), literal: sequence.RenderMarkdown(), children: children);
    }

    private static IReadOnlyList<MarkdownSyntaxNode> BuildInlineLabelChildren(InlineSequence? labelInlines, string fallbackText) {
        if (labelInlines != null && labelInlines.Nodes.Count > 0) {
            return BuildChildren(labelInlines);
        }

        if (string.IsNullOrEmpty(fallbackText)) {
            return Array.Empty<MarkdownSyntaxNode>();
        }

        return new[] {
            new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineText, literal: fallbackText)
        };
    }

    private static IReadOnlyList<MarkdownSyntaxNode> BuildInlineImageChildren(string alt, string src, string? title, string? linkTitle) {
        var nodes = new List<MarkdownSyntaxNode>(4) {
            new MarkdownSyntaxNode(MarkdownSyntaxKind.ImageAlt, literal: alt ?? string.Empty),
            new MarkdownSyntaxNode(MarkdownSyntaxKind.ImageSource, literal: src ?? string.Empty)
        };

        if (!string.IsNullOrEmpty(title)) {
            nodes.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.ImageTitle, literal: title));
        }

        if (!string.IsNullOrEmpty(linkTitle)) {
            nodes.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.ImageLinkTitle, literal: linkTitle));
        }

        return nodes;
    }
}
