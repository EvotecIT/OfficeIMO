using OfficeIMO.Markdown;

namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Renders Markdown to HTML suitable for WebView2/browser hosts, and provides a reusable shell page
/// + an incremental update mechanism.
/// </summary>
public static partial class MarkdownRenderer {
    /// <summary>
    /// Parses Markdown using renderer preprocessing and renderer-owned AST transforms,
    /// returning the final <see cref="MarkdownDoc"/> that would be rendered to HTML.
    /// </summary>
    public static MarkdownDoc ParseDocument(
        string markdown,
        MarkdownRendererOptions? options = null,
        ICollection<MarkdownDocumentTransformDiagnostic>? diagnostics = null,
        ICollection<MarkdownRendererPreProcessorDiagnostic>? preProcessorDiagnostics = null) {
        if (diagnostics == null && preProcessorDiagnostics == null) {
            options ??= new MarkdownRendererOptions();
            var readerOptions = CreateEffectiveReaderOptions(options);
            markdown = PrepareMarkdown(markdown, options, renderErrorAsException: true);
            var doc = MarkdownReader.Parse(markdown, readerOptions);
            return ApplyRendererDocumentTransforms(doc, options, readerOptions, diagnostics: null);
        }

        var result = ParseDocumentResult(markdown, options);
        CopyDiagnostics(result.TransformDiagnostics, diagnostics);
        CopyDiagnostics(result.PreProcessorDiagnostics, preProcessorDiagnostics);
        return result.Document;
    }

    /// <summary>
    /// Parses Markdown using renderer preprocessing and renderer-owned AST transforms,
    /// returning the final document, original syntax tree, and diagnostics together.
    /// </summary>
    public static MarkdownRendererParseResult ParseDocumentResult(
        string markdown,
        MarkdownRendererOptions? options = null) {
        options ??= new MarkdownRendererOptions();
        var readerOptions = CreateEffectiveReaderOptions(options);
        var preProcessorDiagnostics = new List<MarkdownRendererPreProcessorDiagnostic>();

        markdown = PrepareMarkdown(markdown, options, renderErrorAsException: true, preProcessorDiagnostics: preProcessorDiagnostics);

        var parseResult = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics(markdown, readerOptions);
        var transformDiagnostics = new List<MarkdownDocumentTransformDiagnostic>(parseResult.TransformDiagnostics);
        var topLevelBlockSourceSpans = BuildTopLevelBlockSourceSpans(parseResult);
        var readerDiagnosticCount = transformDiagnostics.Count;
        var document = ApplyRendererDocumentTransforms(
            parseResult.Document,
            options,
            readerOptions,
            transformDiagnostics,
            parseResult.SyntaxTree,
            topLevelBlockSourceSpans);
        var rendererDiagnostics = transformDiagnostics.Count > readerDiagnosticCount
            ? transformDiagnostics.Skip(readerDiagnosticCount).ToArray()
            : Array.Empty<MarkdownDocumentTransformDiagnostic>();
        var finalSyntaxTree = rendererDiagnostics.Length == 0
            ? parseResult.FinalSyntaxTree
            : MarkdownReader.BuildFinalSyntaxTree(document, parseResult.FinalSyntaxTree, rendererDiagnostics);

        return new MarkdownRendererParseResult(
            document,
            markdown,
            parseResult.SyntaxTree,
            finalSyntaxTree,
            transformDiagnostics,
            preProcessorDiagnostics);
    }

    /// <summary>
    /// Parses Markdown using the renderer-owned pipeline and returns the AST-backed native projection.
    /// </summary>
    public static MarkdownNativeDocument ParseNativeDocument(
        string markdown,
        MarkdownRendererOptions? options = null) {
        var result = ParseDocumentResult(markdown, options);
        var parseResult = new MarkdownParseResult(
            result.Document,
            result.SyntaxTree,
            result.FinalSyntaxTree,
            result.TransformDiagnostics);
        return MarkdownNativeDocument.FromParseResult(parseResult);
    }

    /// <summary>
    /// Parses Markdown using OfficeIMO.Markdown and returns an HTML fragment (typically an &lt;article class="markdown-body"&gt; wrapper).
    /// When Mermaid is enabled, Mermaid code blocks are annotated with hashes for incremental rendering.
    /// </summary>
    public static string RenderBodyHtml(string markdown, MarkdownRendererOptions? options = null) {
        options ??= new MarkdownRendererOptions();
        var htmlOptions = options.HtmlOptions ?? new HtmlOptions { Kind = HtmlKind.Fragment };
        var readerOptions = CreateEffectiveReaderOptions(options);
        try {
            markdown = PrepareMarkdown(markdown, options, renderErrorAsException: false, htmlOptions);
        } catch (MarkdownPreparationOverflowException ex) {
            return ex.OverflowHtml;
        }
        var doc = MarkdownReader.Parse(markdown, readerOptions);
        doc = ApplyRendererDocumentTransforms(doc, options, readerOptions, diagnostics: null);

        var priorBaseUri = htmlOptions.BaseUri;
        if (!string.IsNullOrWhiteSpace(options.BaseHref) && htmlOptions.BaseUri == null) {
            // Best-effort: use BaseHref for origin restrictions (if enabled). If parsing fails or BaseHref isn't absolute,
            // keep BaseUri null and origin restriction will effectively be disabled.
            if (Uri.TryCreate(options.BaseHref!.Trim(), UriKind.Absolute, out var baseUri)) {
                htmlOptions.BaseUri = baseUri;
            }
        }

        var priorCodeBlockHtmlRenderer = htmlOptions.CodeBlockHtmlRenderer;
        var priorSemanticFencedBlockHtmlRenderer = htmlOptions.SemanticFencedBlockHtmlRenderer;
        htmlOptions.CodeBlockHtmlRenderer = CreateEffectiveCodeBlockHtmlRenderer(options, priorCodeBlockHtmlRenderer);
        htmlOptions.SemanticFencedBlockHtmlRenderer = CreateEffectiveSemanticFencedBlockHtmlRenderer(options, priorSemanticFencedBlockHtmlRenderer);

        string html;
        try {
            html = doc.ToHtmlFragment(htmlOptions) ?? string.Empty;
        } finally {
            htmlOptions.BaseUri = priorBaseUri;
            htmlOptions.CodeBlockHtmlRenderer = priorCodeBlockHtmlRenderer;
            htmlOptions.SemanticFencedBlockHtmlRenderer = priorSemanticFencedBlockHtmlRenderer;
        }

        if (!string.IsNullOrWhiteSpace(options.BaseHref)) {
            // Put <base> into the update payload. The incremental updater moves it into <head>.
            var baseHref = System.Net.WebUtility.HtmlEncode(options.BaseHref!.Trim());
            html = $"<base href=\"{baseHref}\">" + html;
        }

        var post = options.HtmlPostProcessors;
        if (post != null && post.Count > 0) {
            for (int i = 0; i < post.Count; i++) {
                var p = post[i];
                if (p == null) continue;
                html = p(html, options) ?? html ?? string.Empty;
            }
        }

        if (options.MaxBodyHtmlBytes.HasValue && options.MaxBodyHtmlBytes.Value >= 0) {
            int maxBytes = options.MaxBodyHtmlBytes.Value;
            int bytes = Encoding.UTF8.GetByteCount(html ?? string.Empty);
            if (bytes > maxBytes) {
                switch (options.BodyHtmlOverflowHandling) {
                    case OverflowHandling.Throw:
                        throw new InvalidOperationException($"Rendered HTML payload size {bytes} bytes exceeds MaxBodyHtmlBytes {maxBytes}.");
                    case OverflowHandling.Truncate:
                        // Truncating HTML would likely break markup; render an in-band warning instead.
                        return BuildOverflowBodyHtml(htmlOptions, $"Rendered output exceeded the maximum allowed size ({maxBytes} bytes).");
                    case OverflowHandling.RenderError:
                    default:
                        return BuildOverflowBodyHtml(htmlOptions, $"Rendered output exceeded the maximum allowed size ({maxBytes} bytes).");
                }
            }
        }

        return html ?? string.Empty;
    }

}
