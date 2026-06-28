namespace OfficeIMO.Markdown;

internal static class MarkdownInlineSyntaxBuilder {
    private static readonly MarkdownInlineSyntaxBuilderContext _context = new();

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
        if (inline is ISyntaxMarkdownInline syntaxInline) {
            return syntaxInline.BuildSyntaxNode(_context, span);
        }

        switch (inline) {
            case TextRun text:
                var escapedTextChildren = BuildEscapedTextChildren(text);
                return new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.InlineText,
                    span ?? MarkdownBlockSyntaxBuilder.GetAggregateSpan(escapedTextChildren),
                    literal: text.Text,
                    children: escapedTextChildren,
                    associatedObject: text);
            case DecodedHtmlEntityTextRun text:
                var decodedEntityChildren = BuildDecodedEntityChildren(text);
                return new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.InlineText,
                    span ?? MarkdownBlockSyntaxBuilder.GetAggregateSpan(decodedEntityChildren),
                    literal: text.Text,
                    children: decodedEntityChildren,
                    associatedObject: text);
            case CodeSpanInline code:
                var codeSpanChildren = BuildCodeSpanChildren(code);
                return new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.InlineCodeSpan,
                    span ?? MarkdownBlockSyntaxBuilder.GetAggregateSpan(codeSpanChildren),
                    literal: code.Text,
                    children: codeSpanChildren,
                    associatedObject: code);
            case FootnoteRefInline footnote:
                return new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.InlineFootnoteRef,
                    span,
                    literal: footnote.Label,
                    children: BuildFootnoteRefChildren(footnote, span),
                    associatedObject: footnote);
            case HardBreakInline:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineHardBreak, span, literal: "\\n", associatedObject: inline);
            case LinkInline link:
                return new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.InlineLink,
                    span,
                    literal: link.Url,
                    children: BuildInlineLinkChildren(link.LabelInlines, link.Text, link, link.Url, link.Title, link.LinkTarget, link.LinkRel),
                    associatedObject: link);
            case ImageLinkInline imageLink:
                return new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.InlineImageLink,
                    span,
                    literal: imageLink.LinkUrl,
                    children: BuildInlineImageLinkChildren(imageLink),
                    associatedObject: imageLink);
            case ImageInline image:
                return new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.InlineImage,
                    span,
                    literal: image.Src,
                    children: BuildInlineImageChildren(image),
                    associatedObject: image);
            case BoldSequenceInline bold:
                return BuildContainerNode(MarkdownSyntaxKind.InlineStrong, bold, bold.Inlines, span);
            case BoldInline bold:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineStrong, span, literal: bold.Text, associatedObject: bold);
            case ItalicSequenceInline italic:
                return BuildContainerNode(MarkdownSyntaxKind.InlineEmphasis, italic, italic.Inlines, span);
            case ItalicInline italic:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineEmphasis, span, literal: italic.Text, associatedObject: italic);
            case BoldItalicSequenceInline boldItalic:
                return BuildContainerNode(MarkdownSyntaxKind.InlineStrongEmphasis, boldItalic, boldItalic.Inlines, span);
            case BoldItalicInline boldItalic:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineStrongEmphasis, span, literal: boldItalic.Text, associatedObject: boldItalic);
            case StrikethroughSequenceInline strike:
                return BuildContainerNode(MarkdownSyntaxKind.InlineStrikethrough, strike, strike.Inlines, span);
            case StrikethroughInline strike:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineStrikethrough, span, literal: strike.Text, associatedObject: strike);
            case HighlightSequenceInline highlight:
                return BuildContainerNode(MarkdownSyntaxKind.InlineHighlight, highlight, highlight.Inlines, span);
            case HighlightInline highlight:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineHighlight, span, literal: highlight.Text, associatedObject: highlight);
            case UnderlineInline underline:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineUnderline, span, literal: underline.Text, associatedObject: underline);
            case HtmlTagSequenceInline htmlTag:
                var htmlTagChildren = BuildChildren(htmlTag.Inlines);
                var htmlTagMarkerChildren = BuildInlineMarkerChildren(htmlTag, htmlTagChildren);
                return new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.InlineHtmlTag,
                    span ?? MarkdownBlockSyntaxBuilder.GetAggregateSpan(htmlTagMarkerChildren),
                    literal: htmlTag.TagName,
                    children: htmlTagMarkerChildren,
                    associatedObject: htmlTag);
            case HtmlRawInline html:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineHtmlRaw, span, literal: html.Html, associatedObject: html);
            case IInlineContainerMarkdownInline container when container.NestedInlines != null:
                return new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.Unknown,
                    span ?? MarkdownBlockSyntaxBuilder.GetAggregateSpan(BuildChildren(container.NestedInlines)),
                    literal: ((IRenderableMarkdownInline)inline).RenderMarkdown(),
                    children: BuildChildren(container.NestedInlines),
                    associatedObject: inline);
            case IRenderableMarkdownInline renderable:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.Unknown, span, literal: renderable.RenderMarkdown(), associatedObject: inline);
            default:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.Unknown, span, literal: inline.ToString(), associatedObject: inline);
        }
    }

    private static MarkdownSyntaxNode BuildContainerNode(MarkdownSyntaxKind kind, MarkdownInline owner, InlineSequence sequence, MarkdownSourceSpan? span) {
        var children = BuildChildren(sequence);
        var markerChildren = BuildInlineMarkerChildren(owner, children);
        return new MarkdownSyntaxNode(
            kind,
            span ?? MarkdownBlockSyntaxBuilder.GetAggregateSpan(markerChildren),
            literal: sequence.RenderMarkdown(),
            children: markerChildren,
            associatedObject: owner);
    }

    private static IReadOnlyList<MarkdownSyntaxNode> BuildInlineMarkerChildren(MarkdownInline owner, IReadOnlyList<MarkdownSyntaxNode> contentChildren) {
        var openingMarkerSpan = MarkdownInlineMetadataSourceSpans.GetOpeningMarkerSpan(owner);
        var separatorMarkerSpan = MarkdownInlineMetadataSourceSpans.GetSeparatorMarkerSpan(owner);
        var closingMarkerSpan = MarkdownInlineMetadataSourceSpans.GetClosingMarkerSpan(owner);

        if (!openingMarkerSpan.HasValue && !separatorMarkerSpan.HasValue && !closingMarkerSpan.HasValue) {
            return contentChildren;
        }

        var nodes = new List<MarkdownSyntaxNode>(contentChildren.Count + 3);
        AddMarkerNode(
            nodes,
            MarkdownSyntaxKind.InlineOpeningMarker,
            MarkdownInlineMetadataSourceSpans.GetOpeningMarker(owner),
            openingMarkerSpan);

        for (int i = 0; i < contentChildren.Count; i++) {
            nodes.Add(contentChildren[i]);
        }

        AddMarkerNode(
            nodes,
            MarkdownSyntaxKind.InlineSeparatorMarker,
            MarkdownInlineMetadataSourceSpans.GetSeparatorMarker(owner),
            separatorMarkerSpan);
        AddMarkerNode(
            nodes,
            MarkdownSyntaxKind.InlineClosingMarker,
            MarkdownInlineMetadataSourceSpans.GetClosingMarker(owner),
            closingMarkerSpan);

        return nodes;
    }

    private static IReadOnlyList<MarkdownSyntaxNode> BuildEscapedTextChildren(TextRun text) {
        var markerSpan = MarkdownInlineMetadataSourceSpans.GetEscapeMarkerSpan(text);
        var characterSpan = MarkdownInlineMetadataSourceSpans.GetEscapedCharacterSpan(text);

        if (!markerSpan.HasValue && !characterSpan.HasValue) {
            return Array.Empty<MarkdownSyntaxNode>();
        }

        var nodes = new List<MarkdownSyntaxNode>(2);
        AddMarkerNode(
            nodes,
            MarkdownSyntaxKind.InlineEscapeMarker,
            MarkdownInlineMetadataSourceSpans.GetEscapeMarker(text),
            markerSpan);
        AddMarkerNode(
            nodes,
            MarkdownSyntaxKind.InlineEscapedCharacter,
            MarkdownInlineMetadataSourceSpans.GetEscapedCharacter(text),
            characterSpan);

        return nodes;
    }

    private static IReadOnlyList<MarkdownSyntaxNode> BuildDecodedEntityChildren(DecodedHtmlEntityTextRun text) {
        var sourceTextSpan = MarkdownInlineMetadataSourceSpans.GetDecodedEntitySourceTextSpan(text);
        if (!sourceTextSpan.HasValue) {
            return Array.Empty<MarkdownSyntaxNode>();
        }

        return new[] {
            new MarkdownSyntaxNode(
                MarkdownSyntaxKind.InlineEntitySourceText,
                sourceTextSpan,
                literal: MarkdownInlineMetadataSourceSpans.GetDecodedEntitySourceText(text) ?? string.Empty)
        };
    }

    private static IReadOnlyList<MarkdownSyntaxNode> BuildCodeSpanChildren(CodeSpanInline code) {
        var openingMarkerSpan = MarkdownInlineMetadataSourceSpans.GetOpeningMarkerSpan(code);
        var contentSpan = MarkdownInlineMetadataSourceSpans.GetCodeSpanContentSpan(code);
        var closingMarkerSpan = MarkdownInlineMetadataSourceSpans.GetClosingMarkerSpan(code);

        if (!openingMarkerSpan.HasValue && !contentSpan.HasValue && !closingMarkerSpan.HasValue) {
            return Array.Empty<MarkdownSyntaxNode>();
        }

        var nodes = new List<MarkdownSyntaxNode>(3);
        AddMarkerNode(
            nodes,
            MarkdownSyntaxKind.InlineOpeningMarker,
            MarkdownInlineMetadataSourceSpans.GetOpeningMarker(code),
            openingMarkerSpan);

        if (contentSpan.HasValue) {
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.InlineCodeSpanContent,
                contentSpan,
                literal: code.Text));
        }

        AddMarkerNode(
            nodes,
            MarkdownSyntaxKind.InlineClosingMarker,
            MarkdownInlineMetadataSourceSpans.GetClosingMarker(code),
            closingMarkerSpan);

        return nodes;
    }

    private static void AddMarkerNode(List<MarkdownSyntaxNode> nodes, MarkdownSyntaxKind kind, string? marker, MarkdownSourceSpan? span) {
        if (!span.HasValue) {
            return;
        }

        nodes.Add(new MarkdownSyntaxNode(kind, span, literal: marker ?? string.Empty));
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

    private static IReadOnlyList<MarkdownSyntaxNode> BuildInlineLinkChildren(
        InlineSequence? labelInlines,
        string fallbackText,
        LinkInline owner,
        string url,
        string? title,
        string? linkTarget,
        string? linkRel) {
        var nodes = new List<MarkdownSyntaxNode>();
        var labelChildren = BuildInlineLabelChildren(labelInlines, fallbackText);
        for (int i = 0; i < labelChildren.Count; i++) {
            nodes.Add(labelChildren[i]);
        }

        nodes.Add(new MarkdownSyntaxNode(
            MarkdownSyntaxKind.InlineLinkTarget,
            MarkdownInlineMetadataSourceSpans.GetLinkTargetSpan(owner),
            literal: url ?? string.Empty));

        if (!string.IsNullOrEmpty(title)) {
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.InlineLinkTitle,
                MarkdownInlineMetadataSourceSpans.GetLinkTitleSpan(owner),
                literal: title));
        }

        if (!string.IsNullOrEmpty(linkTarget)) {
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.InlineLinkHtmlTarget,
                MarkdownInlineMetadataSourceSpans.GetLinkHtmlTargetSpan(owner),
                literal: linkTarget));
        }

        if (!string.IsNullOrEmpty(linkRel)) {
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.InlineLinkHtmlRel,
                MarkdownInlineMetadataSourceSpans.GetLinkHtmlRelSpan(owner),
                literal: linkRel));
        }

        return nodes;
    }

    private static IReadOnlyList<MarkdownSyntaxNode> BuildFootnoteRefChildren(FootnoteRefInline footnote, MarkdownSourceSpan? span) {
        if (footnote == null || string.IsNullOrEmpty(footnote.Label)) {
            return Array.Empty<MarkdownSyntaxNode>();
        }

        return new[] {
            new MarkdownSyntaxNode(
                MarkdownSyntaxKind.InlineFootnoteLabel,
                GetFootnoteRefLabelSpan(footnote, span),
                literal: footnote.Label)
        };
    }

    private static MarkdownSourceSpan? GetFootnoteRefLabelSpan(FootnoteRefInline footnote, MarkdownSourceSpan? span) {
        if (footnote == null || string.IsNullOrEmpty(footnote.Label) || !span.HasValue || !span.Value.StartColumn.HasValue) {
            return null;
        }

        var labelStartColumn = span.Value.StartColumn.Value + 2;
        var labelEndColumn = labelStartColumn + footnote.Label.Length - 1;
        int? labelStartOffset = null;
        int? labelEndOffset = null;
        if (span.Value.StartOffset.HasValue) {
            labelStartOffset = span.Value.StartOffset.Value + 2;
            labelEndOffset = labelStartOffset.Value + footnote.Label.Length - 1;
        }

        return new MarkdownSourceSpan(
            span.Value.StartLine,
            labelStartColumn,
            span.Value.StartLine,
            labelEndColumn,
            labelStartOffset,
            labelEndOffset);
    }

    private static IReadOnlyList<MarkdownSyntaxNode> BuildInlineImageChildren(ImageInline image) {
        var nodes = new List<MarkdownSyntaxNode>(3) {
            new MarkdownSyntaxNode(
                MarkdownSyntaxKind.ImageAlt,
                MarkdownInlineMetadataSourceSpans.GetImageAltSpan(image),
                literal: image.Alt ?? string.Empty),
            new MarkdownSyntaxNode(
                MarkdownSyntaxKind.ImageSource,
                MarkdownInlineMetadataSourceSpans.GetImageSourceSpan(image),
                literal: image.Src ?? string.Empty)
        };

        if (!string.IsNullOrEmpty(image.Title)) {
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.ImageTitle,
                MarkdownInlineMetadataSourceSpans.GetImageTitleSpan(image),
                literal: image.Title));
        }

        return nodes;
    }

    private static IReadOnlyList<MarkdownSyntaxNode> BuildInlineImageLinkChildren(ImageLinkInline imageLink) {
        var nodes = new List<MarkdownSyntaxNode>(5) {
            new MarkdownSyntaxNode(
                MarkdownSyntaxKind.ImageAlt,
                MarkdownInlineMetadataSourceSpans.GetImageAltSpan(imageLink),
                literal: imageLink.Alt ?? string.Empty),
            new MarkdownSyntaxNode(
                MarkdownSyntaxKind.ImageSource,
                MarkdownInlineMetadataSourceSpans.GetImageSourceSpan(imageLink),
                literal: imageLink.ImageUrl ?? string.Empty),
            new MarkdownSyntaxNode(
                MarkdownSyntaxKind.ImageLinkTarget,
                MarkdownInlineMetadataSourceSpans.GetImageLinkTargetSpan(imageLink),
                literal: imageLink.LinkUrl ?? string.Empty)
        };

        if (!string.IsNullOrEmpty(imageLink.LinkTitle)) {
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.ImageLinkTitle,
                MarkdownInlineMetadataSourceSpans.GetImageLinkTitleSpan(imageLink),
                literal: imageLink.LinkTitle));
        }

        if (!string.IsNullOrEmpty(imageLink.Title)) {
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.ImageTitle,
                MarkdownInlineMetadataSourceSpans.GetImageTitleSpan(imageLink),
                literal: imageLink.Title));
        }

        return nodes;
    }
}
