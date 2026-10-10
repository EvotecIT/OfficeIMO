using System.Threading;

namespace OfficeIMO.Markdown;

public partial class MarkdownDoc {
    /// <summary>Renders the document to Markdown text.</summary>
    public string ToMarkdown() => ToMarkdown(options: null);

    /// <summary>Renders Markdown using optional writer extensions or portability fallbacks.</summary>
    public string ToMarkdown(MarkdownWriteOptions? options) => ToMarkdown(options, CancellationToken.None);

    /// <summary>Renders Markdown with cooperative cancellation between blocks, inline nodes and writer extensions.</summary>
    public string ToMarkdown(MarkdownWriteOptions? options, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        options ??= MarkdownWriteOptions.CreateOfficeIMOProfile();
        var (blocks, headingCatalog) = GetBlocksAndHeadingSlugs(cancellationToken: cancellationToken);
        var context = new MarkdownWriteContext(this, blocks, options, headingCatalog, cancellationToken);
        using var _ctx = MarkdownRenderContext.Push(context);
        var sb = new StringBuilder();
        var renderedFrontMatter = RenderFrontMatter(_frontMatter, options);
        cancellationToken.ThrowIfCancellationRequested();
        if (!string.IsNullOrEmpty(renderedFrontMatter)) {
            sb.AppendLine(renderedFrontMatter);
            sb.AppendLine();
        }
        if (AppendParseOwnedAbbreviationDefinitions(sb, _parseResult, options) && blocks.Count > 0) sb.AppendLine();
        for (int i = 0; i < blocks.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            string rendered = MarkdownBlockRenderDispatcher.RenderMarkdown(blocks[i], context);
            if (!string.IsNullOrEmpty(rendered)) sb.AppendLine(rendered);
            if (i < blocks.Count - 1) sb.AppendLine();
        }
        string markdown = sb.ToString();
        if (options.OutputLineEnding != null) markdown = NormalizeLineEndings(markdown, options.OutputLineEnding);
        cancellationToken.ThrowIfCancellationRequested();
        return markdown;
    }
}
