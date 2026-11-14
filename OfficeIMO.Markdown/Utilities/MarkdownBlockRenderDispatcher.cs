namespace OfficeIMO.Markdown;

internal static class MarkdownBlockRenderDispatcher {
    internal static string RenderHtml(IMarkdownBlock block) {
        if (block == null) {
            return string.Empty;
        }

        System.Runtime.CompilerServices.RuntimeHelpers.EnsureSufficientExecutionStack();

        var context = HtmlRenderContext.BodyContext;
        return context == null
            ? block.RenderHtml()
            : RenderHtml(block, context);
    }

    internal static string RenderHtml(IMarkdownBlock block, MarkdownBodyRenderContext context) {
        if (block == null) {
            return string.Empty;
        }

        // Container blocks render their children through here; a hand-built tree deeper than the
        // stack fails with a catchable exception instead of ending the process.
        System.Runtime.CompilerServices.RuntimeHelpers.EnsureSufficientExecutionStack();

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

    /// <summary>Appends the same HTML as <see cref="RenderHtml(IMarkdownBlock)"/> using the ambient body context.</summary>
    internal static void AppendHtml(StringBuilder builder, IMarkdownBlock block) {
        if (block == null) {
            return;
        }

        var context = HtmlRenderContext.BodyContext;
        if (context == null) {
            System.Runtime.CompilerServices.RuntimeHelpers.EnsureSufficientExecutionStack();
            AppendDefaultHtml(builder, block);
            return;
        }

        AppendHtml(builder, block, context);
    }

    /// <summary>Appends the same HTML as <see cref="RenderHtml(IMarkdownBlock, MarkdownBodyRenderContext)"/>.</summary>
    internal static void AppendHtml(StringBuilder builder, IMarkdownBlock block, MarkdownBodyRenderContext context) {
        if (block == null) {
            return;
        }

        System.Runtime.CompilerServices.RuntimeHelpers.EnsureSufficientExecutionStack();

        var overridden = TryRenderSyntaxBlockHtmlOverride(block, context);
        if (overridden != null) {
            builder.Append(overridden);
            return;
        }

        overridden = TryRenderBlockHtmlOverride(block, context);
        if (overridden != null) {
            builder.Append(overridden);
            return;
        }

        if (block is IContextualHtmlMarkdownBlock contextualBlock) {
            if (block is IHtmlAppendableContextualMarkdownBlock appendable) {
                appendable.AppendHtml(builder, context);
            } else {
                builder.Append(contextualBlock.RenderHtml(context));
            }

            return;
        }

        AppendDefaultHtml(builder, block);
    }

    private static void AppendDefaultHtml(StringBuilder builder, IMarkdownBlock block) {
        if (block is IHtmlAppendableMarkdownBlock appendable) {
            appendable.AppendHtml(builder);
            return;
        }

        builder.Append(block.RenderHtml());
    }

    internal static string RenderTightListItemHtml(IMarkdownBlock block) {
        if (block == null) {
            return string.Empty;
        }

        var context = HtmlRenderContext.BodyContext;
        return context == null
            ? RenderTightListItemHtmlFallback(block)
            : RenderTightListItemHtml(block, context);
    }

    internal static string RenderTightListItemHtml(IMarkdownBlock block, MarkdownBodyRenderContext context) {
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

        return RenderTightListItemHtmlFallback(block, context);
    }

    internal static string RenderMarkdown(IMarkdownBlock block) {
        if (block == null) {
            return string.Empty;
        }

        System.Runtime.CompilerServices.RuntimeHelpers.EnsureSufficientExecutionStack();

        var context = MarkdownRenderContext.WriteContext;
        return context == null
            ? block.RenderMarkdown()
            : RenderMarkdown(block, context);
    }

    internal static string RenderMarkdown(IMarkdownBlock block, MarkdownWriteContext context) {
        context.CancellationToken.ThrowIfCancellationRequested();
        if (block == null) {
            return string.Empty;
        }

        var overridden = TryRenderSyntaxBlockMarkdownOverride(block, context);
        if (overridden != null) {
            context.CancellationToken.ThrowIfCancellationRequested();
            return overridden;
        }

        overridden = TryRenderBlockMarkdownOverride(block, context);
        if (overridden != null) {
            context.CancellationToken.ThrowIfCancellationRequested();
            return overridden;
        }

        context.CancellationToken.ThrowIfCancellationRequested();
        string rendered = block.RenderMarkdown();
        context.CancellationToken.ThrowIfCancellationRequested();
        return rendered;
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
            context.CancellationToken.ThrowIfCancellationRequested();
            var extension = extensions[i];
            if (extension == null || !extension.Matches(syntaxNode)) {
                continue;
            }

            var rendered = extension.RenderMarkdown(block, syntaxNode, context);
            context.CancellationToken.ThrowIfCancellationRequested();
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
            context.CancellationToken.ThrowIfCancellationRequested();
            var extension = extensions[i];
            if (extension == null || !extension.Matches(block)) {
                continue;
            }

            var rendered = extension.RenderMarkdown(block, context);
            context.CancellationToken.ThrowIfCancellationRequested();
            if (rendered != null) {
                return rendered;
            }
        }

        return null;
    }

    private static string RenderTightListItemHtmlFallback(IMarkdownBlock block, MarkdownBodyRenderContext? context = null) {
        if (block is ITightListItemHtmlMarkdownBlock tightHtmlBlock) {
            return tightHtmlBlock.RenderTightListItemHtml();
        }

        if (context != null && block is IContextualHtmlMarkdownBlock contextualBlock) {
            return contextualBlock.RenderHtml(context);
        }

        return block.RenderHtml();
    }
}
