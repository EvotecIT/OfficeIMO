namespace OfficeIMO.Markdown;

internal static class MarkdownBlockSyntaxBuilder {
    internal static MarkdownSyntaxNode BuildBlock(IMarkdownBlock block, MarkdownSourceSpan? span = null) {
        if (block is ISyntaxMarkdownBlock syntaxBlock) {
            return syntaxBlock.BuildSyntaxNode(span);
        }

        return new MarkdownSyntaxNode(MarkdownSyntaxKind.Unknown, span, block.RenderMarkdown());
    }

    internal static MarkdownSyntaxNode BuildInlineBlock(IInlineSyntaxMarkdownBlock inlineBlock, MarkdownSourceSpan? span = null) =>
        new MarkdownSyntaxNode(
            inlineBlock.SyntaxKind,
            span ?? inlineBlock.ProvidedSyntaxSpan,
            inlineBlock.SyntaxInlines.RenderMarkdown());

    internal static MarkdownSyntaxNode BuildHeadingBlock(HeadingBlock heading, MarkdownSourceSpan? span) =>
        new MarkdownSyntaxNode(MarkdownSyntaxKind.Heading, span, heading.Inlines.RenderMarkdown(), BuildHeadingChildren(heading, span));

    internal static MarkdownSyntaxNode BuildHorizontalRuleBlock(MarkdownSourceSpan? span) =>
        new MarkdownSyntaxNode(MarkdownSyntaxKind.HorizontalRule, span, "---");

    internal static MarkdownSyntaxNode BuildCodeBlock(CodeBlock code, MarkdownSourceSpan? span) =>
        new MarkdownSyntaxNode(
            MarkdownSyntaxKind.CodeBlock,
            span,
            NormalizeSyntaxLiteralLineEndings(code.Content),
            BuildCodeBlockChildren(code, span));

    internal static MarkdownSyntaxNode BuildImageBlock(ImageBlock image, MarkdownSourceSpan? span) =>
        new MarkdownSyntaxNode(
            MarkdownSyntaxKind.Image,
            span,
            ((IMarkdownBlock)image).RenderMarkdown(),
            BuildImageChildren(image, span));

    internal static MarkdownSyntaxNode BuildTableBlock(TableBlock table, MarkdownSourceSpan? span) =>
        new MarkdownSyntaxNode(
            MarkdownSyntaxKind.Table,
            span,
            ((IMarkdownBlock)table).RenderMarkdown(),
            table.BuildSyntaxChildren(span));

    internal static MarkdownSyntaxNode BuildQuoteBlock(QuoteBlock quote, MarkdownSourceSpan? span) =>
        new MarkdownSyntaxNode(
            MarkdownSyntaxKind.Quote,
            span,
            quote.Children.Count == 0 ? string.Join("\n", quote.Lines) : null,
            GetOwnedSyntaxChildrenOrBuild(quote));

    internal static MarkdownSyntaxNode BuildDefinitionListBlock(DefinitionListBlock definitionList, MarkdownSourceSpan? span) =>
        new MarkdownSyntaxNode(MarkdownSyntaxKind.DefinitionList, span, children: definitionList.BuildSyntaxItems());

    internal static MarkdownSyntaxNode BuildCalloutBlock(CalloutBlock callout, MarkdownSourceSpan? span) {
        var calloutTitleMarkdown = callout.TitleInlines.RenderMarkdown();
        return new MarkdownSyntaxNode(
            MarkdownSyntaxKind.Callout,
            span,
            string.IsNullOrWhiteSpace(calloutTitleMarkdown) ? callout.Kind : callout.Kind + ":" + calloutTitleMarkdown,
            GetOwnedSyntaxChildrenOrBuild(callout));
    }

    internal static MarkdownSyntaxNode BuildDetailsBlock(DetailsBlock details, MarkdownSourceSpan? span) =>
        new MarkdownSyntaxNode(MarkdownSyntaxKind.Details, span, details.Open ? "open" : null, GetOwnedSyntaxChildrenOrBuild(details));

    internal static MarkdownSyntaxNode BuildFootnoteBlock(FootnoteDefinitionBlock footnote, MarkdownSourceSpan? span) =>
        new MarkdownSyntaxNode(
            MarkdownSyntaxKind.FootnoteDefinition,
            span,
            footnote.Label,
            GetOwnedSyntaxChildrenOrBuild(footnote));

    internal static MarkdownSyntaxNode BuildFrontMatterBlock(FrontMatterBlock frontMatter, MarkdownSourceSpan? span) =>
        new MarkdownSyntaxNode(MarkdownSyntaxKind.FrontMatter, span, frontMatter.Render());

    internal static MarkdownSyntaxNode BuildHtmlCommentBlock(HtmlCommentBlock comment, MarkdownSourceSpan? span) =>
        new MarkdownSyntaxNode(MarkdownSyntaxKind.HtmlComment, span, comment.Comment);

    internal static MarkdownSyntaxNode BuildHtmlRawBlock(HtmlRawBlock rawHtml, MarkdownSourceSpan? span) =>
        new MarkdownSyntaxNode(MarkdownSyntaxKind.HtmlRaw, span, rawHtml.Html);

    internal static MarkdownSyntaxNode BuildTocBlock(TocBlock toc, MarkdownSourceSpan? span) =>
        new MarkdownSyntaxNode(MarkdownSyntaxKind.Toc, span, ((IMarkdownBlock)toc).RenderMarkdown());

    internal static MarkdownSyntaxNode BuildTocPlaceholderBlock(MarkdownSourceSpan? span) =>
        new MarkdownSyntaxNode(MarkdownSyntaxKind.TocPlaceholder, span);

    internal static IReadOnlyList<MarkdownSyntaxNode> BuildChildSyntaxNodes(IEnumerable<IMarkdownBlock> children) {
        var nodes = new List<MarkdownSyntaxNode>();
        foreach (var child in children) {
            if (child == null) continue;
            nodes.Add(BuildBlock(child));
        }
        return nodes;
    }

    internal static IReadOnlyList<MarkdownSyntaxNode> GetOwnedSyntaxChildrenOrBuild(IChildMarkdownBlockContainer block) {
        if (block is IOwnedSyntaxChildrenMarkdownBlock ownedSyntaxChildren) {
            return ownedSyntaxChildren.BuildOwnedSyntaxChildren();
        }

        if (block is ISyntaxChildrenMarkdownBlock syntaxOwner &&
            syntaxOwner.ProvidedSyntaxChildren != null &&
            syntaxOwner.ProvidedSyntaxChildren.Count > 0) {
            return syntaxOwner.ProvidedSyntaxChildren;
        }

        return BuildChildSyntaxNodes(block.ChildBlocks);
    }

    internal static IReadOnlyList<MarkdownSyntaxNode> BuildHeadingChildren(HeadingBlock heading, MarkdownSourceSpan? span) {
        var nodes = new List<MarkdownSyntaxNode> {
            new MarkdownSyntaxNode(MarkdownSyntaxKind.HeadingLevel, literal: heading.Level.ToString(System.Globalization.CultureInfo.InvariantCulture))
        };

        MarkdownSourceSpan? textSpan = span.HasValue ? new MarkdownSourceSpan(span.Value.StartLine, span.Value.StartLine) : null;
        nodes.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.HeadingText, textSpan, heading.Inlines.RenderMarkdown()));

        return nodes;
    }

    internal static IReadOnlyList<MarkdownSyntaxNode> BuildCodeBlockChildren(CodeBlock code, MarkdownSourceSpan? span) {
        if (!span.HasValue) return Array.Empty<MarkdownSyntaxNode>();

        var nodes = new List<MarkdownSyntaxNode>();
        if (code.IsFenced && !string.IsNullOrEmpty(code.Language)) {
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.CodeFenceInfo,
                new MarkdownSourceSpan(span.Value.StartLine, span.Value.StartLine),
                code.Language));
        }

        MarkdownSourceSpan? contentSpan;
        if (code.IsFenced) {
            contentSpan = span.Value.EndLine > span.Value.StartLine + 1
                ? new MarkdownSourceSpan(span.Value.StartLine + 1, span.Value.EndLine - 1)
                : null;
        } else {
            contentSpan = span.Value;
        }

        nodes.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.CodeContent, contentSpan, NormalizeSyntaxLiteralLineEndings(code.Content)));
        return nodes;
    }

    internal static IReadOnlyList<MarkdownSyntaxNode> BuildImageChildren(ImageBlock image, MarkdownSourceSpan? span) {
        if (!span.HasValue) return Array.Empty<MarkdownSyntaxNode>();

        var nodes = new List<MarkdownSyntaxNode> {
            new MarkdownSyntaxNode(MarkdownSyntaxKind.ImageSource, span, image.Path)
        };

        if (!string.IsNullOrEmpty(image.Alt)) {
            nodes.Insert(0, new MarkdownSyntaxNode(MarkdownSyntaxKind.ImageAlt, span, image.Alt));
        }

        if (!string.IsNullOrEmpty(image.Title)) {
            nodes.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.ImageTitle, span, image.Title));
        }

        return nodes;
    }

    internal static string NormalizeSyntaxLiteralLineEndings(string? value) {
        if (string.IsNullOrEmpty(value)) return string.Empty;
        string normalized = value!;
        return normalized.Replace("\r\n", "\n").Replace('\r', '\n');
    }

    internal static MarkdownSourceSpan? GetAggregateSpan(IReadOnlyList<MarkdownSyntaxNode> nodes) {
        if (nodes == null || nodes.Count == 0) return null;

        int? start = null;
        int? end = null;
        for (int i = 0; i < nodes.Count; i++) {
            var span = nodes[i].SourceSpan;
            if (!span.HasValue) continue;

            if (!start.HasValue || span.Value.StartLine < start.Value) start = span.Value.StartLine;
            if (!end.HasValue || span.Value.EndLine > end.Value) end = span.Value.EndLine;
        }

        if (!start.HasValue || !end.HasValue) return null;
        return new MarkdownSourceSpan(start.Value, end.Value);
    }
}
