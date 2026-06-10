namespace OfficeIMO.Markdown;

internal static class MarkdownNativeProjectionFactory {
    internal static MarkdownNativeBlock? Create(
        MarkdownSyntaxNode syntaxNode,
        ICollection<MarkdownNativeDiagnostic> diagnostics) {
        if (syntaxNode?.AssociatedObject is not IMarkdownBlock block) {
            return null;
        }

        MarkdownNativeBlock nativeBlock = block switch {
            HeadingBlock heading => new MarkdownNativeHeadingBlock(heading, syntaxNode),
            ParagraphBlock paragraph => new MarkdownNativeParagraphBlock(paragraph, syntaxNode),
            OrderedListBlock ordered => new MarkdownNativeListBlock(ordered, syntaxNode, CreateListItems(syntaxNode, diagnostics)),
            UnorderedListBlock unordered => new MarkdownNativeListBlock(unordered, syntaxNode, CreateListItems(syntaxNode, diagnostics)),
            QuoteBlock quote => new MarkdownNativeQuoteBlock(quote, syntaxNode, CreateChildren(syntaxNode, diagnostics)),
            CalloutBlock callout => new MarkdownNativeCalloutBlock(callout, syntaxNode, CreateChildren(syntaxNode, diagnostics)),
            ImageBlock image => new MarkdownNativeImageBlock(image, syntaxNode),
            CodeBlock code => new MarkdownNativeCodeBlock(code, syntaxNode),
            SemanticFencedBlock visual => new MarkdownNativeVisualBlock(visual, syntaxNode),
            TableBlock table => new MarkdownNativeTableBlock(table, syntaxNode, diagnostics),
            DetailsBlock details => new MarkdownNativeDetailsBlock(details, syntaxNode, CreateChildren(syntaxNode, diagnostics, static node => node.AssociatedObject is not SummaryBlock)),
            FrontMatterBlock frontMatter => new MarkdownNativeFrontMatterBlock(frontMatter, syntaxNode),
            HtmlRawBlock html => new MarkdownNativeHtmlBlock(html, syntaxNode),
            HtmlCommentBlock comment => new MarkdownNativeHtmlBlock(comment, syntaxNode),
            _ => new MarkdownNativeOtherBlock(block, syntaxNode)
        };

        if (nativeBlock is MarkdownNativeOtherBlock) {
            diagnostics.Add(new MarkdownNativeDiagnostic(
                "native.unsupported-block",
                $"No specialized native projection exists for markdown block '{block.GetType().Name}'.",
                MarkdownNativeDiagnosticSeverity.Info,
                nativeBlock.SourceSpan,
                nativeBlock));
        }

        return nativeBlock;
    }

    internal static IReadOnlyList<MarkdownNativeBlock> CreateChildren(
        MarkdownSyntaxNode parent,
        ICollection<MarkdownNativeDiagnostic> diagnostics,
        Func<MarkdownSyntaxNode, bool>? include = null) {
        if (parent.Children.Count == 0) {
            return Array.Empty<MarkdownNativeBlock>();
        }

        var children = new List<MarkdownNativeBlock>();
        for (var i = 0; i < parent.Children.Count; i++) {
            var child = parent.Children[i];
            if (include != null && !include(child)) {
                continue;
            }

            var nativeChild = Create(child, diagnostics);
            if (nativeChild != null) {
                children.Add(nativeChild);
            }
        }

        return children;
    }

    private static IReadOnlyList<MarkdownNativeListItem> CreateListItems(
        MarkdownSyntaxNode listNode,
        ICollection<MarkdownNativeDiagnostic> diagnostics) {
        if (listNode.Children.Count == 0) {
            return Array.Empty<MarkdownNativeListItem>();
        }

        var items = new List<MarkdownNativeListItem>();
        for (var i = 0; i < listNode.Children.Count; i++) {
            var child = listNode.Children[i];
            if (child.AssociatedObject is not ListItem item) {
                continue;
            }

            items.Add(new MarkdownNativeListItem(item, child, CreateChildren(child, diagnostics)));
        }

        return items;
    }
}
