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
                    associatedObject: code,
                    attributes: code.Attributes);
            case FootnoteRefInline footnote:
                return new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.InlineFootnoteRef,
                    span,
                    literal: footnote.Label,
                    children: BuildFootnoteRefChildren(footnote, span),
                    associatedObject: footnote);
            case AbbreviationInline abbreviation:
                return new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.InlineAbbreviation,
                    span,
                    literal: abbreviation.Text,
                    children: BuildAbbreviationChildren(abbreviation, span),
                    associatedObject: abbreviation);
            case HardBreakInline hardBreak:
                var hardBreakChildren = BuildHardBreakChildren(hardBreak);
                return new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.InlineHardBreak,
                    span ?? MarkdownBlockSyntaxBuilder.GetAggregateSpan(hardBreakChildren),
                    literal: "\\n",
                    children: hardBreakChildren,
                    associatedObject: inline);
            case LinkInline link:
                return new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.InlineLink,
                    span,
                    literal: link.Url,
                    children: BuildInlineLinkChildren(link.LabelInlines, link.Text, link, link.Url, link.Title, link.LinkTarget, link.LinkRel),
                    associatedObject: link,
                    attributes: link.Attributes);
            case ImageLinkInline imageLink:
                return new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.InlineImageLink,
                    span,
                    literal: imageLink.LinkUrl,
                    children: BuildInlineImageLinkChildren(imageLink),
                    associatedObject: imageLink,
                    attributes: imageLink.Attributes);
            case ImageInline image:
                return new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.InlineImage,
                    span,
                    literal: image.Src,
                    children: BuildInlineImageChildren(image),
                    associatedObject: image,
                    attributes: image.Attributes);
            case BoldSequenceInline bold:
                return BuildContainerNode(MarkdownSyntaxKind.InlineStrong, bold, bold.Inlines, span);
            case BoldInline bold:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineStrong, span, literal: bold.Text, associatedObject: bold, attributes: bold.Attributes);
            case ItalicSequenceInline italic:
                return BuildContainerNode(MarkdownSyntaxKind.InlineEmphasis, italic, italic.Inlines, span);
            case ItalicInline italic:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineEmphasis, span, literal: italic.Text, associatedObject: italic, attributes: italic.Attributes);
            case BoldItalicSequenceInline boldItalic:
                return BuildContainerNode(MarkdownSyntaxKind.InlineStrongEmphasis, boldItalic, boldItalic.Inlines, span);
            case BoldItalicInline boldItalic:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineStrongEmphasis, span, literal: boldItalic.Text, associatedObject: boldItalic, attributes: boldItalic.Attributes);
            case StrikethroughSequenceInline strike:
                return BuildContainerNode(MarkdownSyntaxKind.InlineStrikethrough, strike, strike.Inlines, span);
            case StrikethroughInline strike:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineStrikethrough, span, literal: strike.Text, associatedObject: strike, attributes: strike.Attributes);
            case HighlightSequenceInline highlight:
                return BuildContainerNode(MarkdownSyntaxKind.InlineHighlight, highlight, highlight.Inlines, span);
            case HighlightInline highlight:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineHighlight, span, literal: highlight.Text, associatedObject: highlight, attributes: highlight.Attributes);
            case InsertedSequenceInline inserted:
                return BuildContainerNode(MarkdownSyntaxKind.InlineInserted, inserted, inserted.Inlines, span);
            case InsertedInline inserted:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineInserted, span, literal: inserted.Text, associatedObject: inserted, attributes: inserted.Attributes);
            case SuperscriptSequenceInline superscript:
                return BuildContainerNode(MarkdownSyntaxKind.InlineSuperscript, superscript, superscript.Inlines, span);
            case SuperscriptInline superscript:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineSuperscript, span, literal: superscript.Text, associatedObject: superscript, attributes: superscript.Attributes);
            case SubscriptSequenceInline subscript:
                return BuildContainerNode(MarkdownSyntaxKind.InlineSubscript, subscript, subscript.Inlines, span);
            case SubscriptInline subscript:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineSubscript, span, literal: subscript.Text, associatedObject: subscript, attributes: subscript.Attributes);
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
            associatedObject: owner,
            attributes: owner.Attributes);
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

    private static IReadOnlyList<MarkdownSyntaxNode> BuildHardBreakChildren(HardBreakInline hardBreak) {
        var markerSpan = MarkdownInlineMetadataSourceSpans.GetHardBreakMarkerSpan(hardBreak);
        if (!markerSpan.HasValue) {
            return Array.Empty<MarkdownSyntaxNode>();
        }

        return new[] {
            new MarkdownSyntaxNode(
                MarkdownSyntaxKind.InlineHardBreakMarker,
                markerSpan,
                literal: MarkdownInlineMetadataSourceSpans.GetHardBreakMarker(hardBreak) ?? string.Empty)
        };
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
        AddMarkerNode(
            nodes,
            MarkdownSyntaxKind.InlineOpeningMarker,
            MarkdownInlineMetadataSourceSpans.GetOpeningMarker(owner),
            MarkdownInlineMetadataSourceSpans.GetOpeningMarkerSpan(owner));

        var labelChildren = BuildInlineLabelChildren(labelInlines, fallbackText);
        for (int i = 0; i < labelChildren.Count; i++) {
            nodes.Add(labelChildren[i]);
        }

        AddMarkerNode(
            nodes,
            MarkdownSyntaxKind.InlineSeparatorMarker,
            MarkdownInlineMetadataSourceSpans.GetSeparatorMarker(owner),
            MarkdownInlineMetadataSourceSpans.GetSeparatorMarkerSpan(owner));

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

        AddMarkerNode(
            nodes,
            MarkdownSyntaxKind.InlineClosingMarker,
            MarkdownInlineMetadataSourceSpans.GetClosingMarker(owner),
            MarkdownInlineMetadataSourceSpans.GetClosingMarkerSpan(owner));

        return nodes;
    }

    private static IReadOnlyList<MarkdownSyntaxNode> BuildFootnoteRefChildren(FootnoteRefInline footnote, MarkdownSourceSpan? span) {
        if (footnote == null || string.IsNullOrEmpty(footnote.Label)) {
            return Array.Empty<MarkdownSyntaxNode>();
        }

        var nodes = new List<MarkdownSyntaxNode>(3);
        AddMarkerNode(
            nodes,
            MarkdownSyntaxKind.InlineOpeningMarker,
            MarkdownInlineMetadataSourceSpans.GetOpeningMarker(footnote),
            MarkdownInlineMetadataSourceSpans.GetOpeningMarkerSpan(footnote));

        nodes.Add(new MarkdownSyntaxNode(
            MarkdownSyntaxKind.InlineFootnoteLabel,
            GetFootnoteRefLabelSpan(footnote, span),
            literal: footnote.Label));

        AddMarkerNode(
            nodes,
            MarkdownSyntaxKind.InlineClosingMarker,
            MarkdownInlineMetadataSourceSpans.GetClosingMarker(footnote),
            MarkdownInlineMetadataSourceSpans.GetClosingMarkerSpan(footnote));

        return nodes;
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

    private static IReadOnlyList<MarkdownSyntaxNode> BuildAbbreviationChildren(AbbreviationInline abbreviation, MarkdownSourceSpan? span) {
        if (abbreviation == null) {
            return Array.Empty<MarkdownSyntaxNode>();
        }

        var nodes = new List<MarkdownSyntaxNode>(2) {
            new MarkdownSyntaxNode(
                MarkdownSyntaxKind.InlineAbbreviationText,
                MarkdownInlineMetadataSourceSpans.GetAbbreviationTextSpan(abbreviation) ?? span,
                literal: abbreviation.Text)
        };

        nodes.Add(new MarkdownSyntaxNode(
            MarkdownSyntaxKind.InlineAbbreviationTitle,
            MarkdownInlineMetadataSourceSpans.GetAbbreviationTitleSpan(abbreviation),
            literal: abbreviation.Title));
        return nodes;
    }

    private static IReadOnlyList<MarkdownSyntaxNode> BuildInlineImageChildren(ImageInline image) {
        var nodes = new List<MarkdownSyntaxNode>(6);
        AddMarkerNode(
            nodes,
            MarkdownSyntaxKind.InlineOpeningMarker,
            MarkdownInlineMetadataSourceSpans.GetOpeningMarker(image),
            MarkdownInlineMetadataSourceSpans.GetOpeningMarkerSpan(image));

        nodes.Add(new MarkdownSyntaxNode(
            MarkdownSyntaxKind.ImageAlt,
            MarkdownInlineMetadataSourceSpans.GetImageAltSpan(image),
            literal: image.Alt ?? string.Empty));

        AddMarkerNode(
            nodes,
            MarkdownSyntaxKind.InlineSeparatorMarker,
            MarkdownInlineMetadataSourceSpans.GetSeparatorMarker(image),
            MarkdownInlineMetadataSourceSpans.GetSeparatorMarkerSpan(image));

        nodes.Add(new MarkdownSyntaxNode(
            MarkdownSyntaxKind.ImageSource,
            MarkdownInlineMetadataSourceSpans.GetImageSourceSpan(image),
            literal: image.Src ?? string.Empty));

        if (!string.IsNullOrEmpty(image.Title)) {
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.ImageTitle,
                MarkdownInlineMetadataSourceSpans.GetImageTitleSpan(image),
                literal: image.Title));
        }

        AddMarkerNode(
            nodes,
            MarkdownSyntaxKind.InlineClosingMarker,
            MarkdownInlineMetadataSourceSpans.GetClosingMarker(image),
            MarkdownInlineMetadataSourceSpans.GetClosingMarkerSpan(image));

        return nodes;
    }

    private static IReadOnlyList<MarkdownSyntaxNode> BuildInlineImageLinkChildren(ImageLinkInline imageLink) {
        var nodes = new List<MarkdownSyntaxNode>(8);
        AddMarkerNode(
            nodes,
            MarkdownSyntaxKind.InlineOpeningMarker,
            MarkdownInlineMetadataSourceSpans.GetOpeningMarker(imageLink),
            MarkdownInlineMetadataSourceSpans.GetOpeningMarkerSpan(imageLink));

        nodes.Add(new MarkdownSyntaxNode(
            MarkdownSyntaxKind.ImageAlt,
            MarkdownInlineMetadataSourceSpans.GetImageAltSpan(imageLink),
            literal: imageLink.Alt ?? string.Empty));
        nodes.Add(new MarkdownSyntaxNode(
            MarkdownSyntaxKind.ImageSource,
            MarkdownInlineMetadataSourceSpans.GetImageSourceSpan(imageLink),
            literal: imageLink.ImageUrl ?? string.Empty));

        if (!string.IsNullOrEmpty(imageLink.Title)) {
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.ImageTitle,
                MarkdownInlineMetadataSourceSpans.GetImageTitleSpan(imageLink),
                literal: imageLink.Title));
        }

        AddMarkerNode(
            nodes,
            MarkdownSyntaxKind.InlineSeparatorMarker,
            MarkdownInlineMetadataSourceSpans.GetSeparatorMarker(imageLink),
            MarkdownInlineMetadataSourceSpans.GetSeparatorMarkerSpan(imageLink));

        nodes.Add(new MarkdownSyntaxNode(
            MarkdownSyntaxKind.ImageLinkTarget,
            MarkdownInlineMetadataSourceSpans.GetImageLinkTargetSpan(imageLink),
            literal: imageLink.LinkUrl ?? string.Empty));

        if (!string.IsNullOrEmpty(imageLink.LinkTitle)) {
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.ImageLinkTitle,
                MarkdownInlineMetadataSourceSpans.GetImageLinkTitleSpan(imageLink),
                literal: imageLink.LinkTitle));
        }

        AddMarkerNode(
            nodes,
            MarkdownSyntaxKind.InlineClosingMarker,
            MarkdownInlineMetadataSourceSpans.GetClosingMarker(imageLink),
            MarkdownInlineMetadataSourceSpans.GetClosingMarkerSpan(imageLink));

        return nodes;
    }
}
