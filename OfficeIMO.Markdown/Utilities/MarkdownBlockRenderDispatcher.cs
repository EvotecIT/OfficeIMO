namespace OfficeIMO.Markdown;

internal static class MarkdownBlockRenderDispatcher {
    internal static string RenderHtml(IMarkdownBlock block) {
        if (block == null) {
            return string.Empty;
        }

        var context = HtmlRenderContext.BodyContext;
        return context == null
            ? block.RenderHtml()
            : RenderHtml(block, context);
    }

    internal static string RenderHtml(IMarkdownBlock block, MarkdownBodyRenderContext context) {
        if (block == null) {
            return string.Empty;
        }

        var overridden = TryRenderSyntaxBlockHtmlOverride(block, context);
        if (overridden != null) {
            return overridden;
        }

        overridden = TryRenderBlockHtmlOverride(block, context);
        if (overridden != null) {
            return overridden;
        }

        if (block is IContextualHtmlMarkdownBlock contextualBlock) {
            return contextualBlock.RenderHtml(context);
        }

        return block.RenderHtml();
    }

    internal static string RenderMarkdown(IMarkdownBlock block) {
        if (block == null) {
            return string.Empty;
        }

        var context = MarkdownRenderContext.WriteContext;
        return context == null
            ? block.RenderMarkdown()
            : RenderMarkdown(block, context);
    }

    internal static string RenderMarkdown(IMarkdownBlock block, MarkdownWriteContext context) {
        if (block == null) {
            return string.Empty;
        }

        var overridden = TryRenderSyntaxBlockMarkdownOverride(block, context);
        if (overridden != null) {
            return overridden;
        }

        overridden = TryRenderBlockMarkdownOverride(block, context);
        if (overridden != null) {
            return overridden;
        }

        return block.RenderMarkdown();
    }

    private static string? TryRenderSyntaxBlockHtmlOverride(IMarkdownBlock block, MarkdownBodyRenderContext context) {
        var extensions = context.Options.SyntaxBlockRenderExtensions;
        if (extensions.Count == 0) {
            return null;
        }

        var syntaxNode = context.FindSyntaxNode(block);
        if (syntaxNode == null) {
            return null;
        }

        for (int i = extensions.Count - 1; i >= 0; i--) {
            var extension = extensions[i];
            if (extension == null || !extension.Matches(syntaxNode)) {
                continue;
            }

            var rendered = extension.RenderHtml(block, syntaxNode, context);
            if (rendered != null) {
                return rendered;
            }
        }

        return null;
    }

    private static string? TryRenderBlockHtmlOverride(IMarkdownBlock block, MarkdownBodyRenderContext context) {
        var extensions = context.Options.BlockRenderExtensions;
        if (extensions.Count == 0) {
            return null;
        }

        for (int i = extensions.Count - 1; i >= 0; i--) {
            var extension = extensions[i];
            if (extension == null || !extension.Matches(block)) {
                continue;
            }

            var rendered = extension.RenderHtml(block, context);
            if (rendered != null) {
                return rendered;
            }
        }

        return null;
    }

    private static string? TryRenderSyntaxBlockMarkdownOverride(IMarkdownBlock block, MarkdownWriteContext context) {
        var extensions = context.Options.SyntaxBlockRenderExtensions;
        if (extensions.Count == 0) {
            return null;
        }

        var syntaxNode = context.FindSyntaxNode(block);
        if (syntaxNode == null) {
            return null;
        }

        for (int i = extensions.Count - 1; i >= 0; i--) {
            var extension = extensions[i];
            if (extension == null || !extension.Matches(syntaxNode)) {
                continue;
            }

            var rendered = extension.RenderMarkdown(block, syntaxNode, context);
            if (rendered != null) {
                return rendered;
            }
        }

        return null;
    }

    private static string? TryRenderBlockMarkdownOverride(IMarkdownBlock block, MarkdownWriteContext context) {
        var extensions = context.Options.BlockRenderExtensions;
        if (extensions.Count == 0) {
            return null;
        }

        for (int i = extensions.Count - 1; i >= 0; i--) {
            var extension = extensions[i];
            if (extension == null || !extension.Matches(block)) {
                continue;
            }

            var rendered = extension.RenderMarkdown(block, context);
            if (rendered != null) {
                return rendered;
            }
        }

        return null;
    }
}
