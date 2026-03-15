using System.Security.Cryptography;
using OfficeIMO.Markdown;

namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Renders Markdown to HTML suitable for WebView2/browser hosts, and provides a reusable shell page
/// + an incremental update mechanism.
/// </summary>
public static class MarkdownRenderer {
    /// <summary>
    /// Parses Markdown using OfficeIMO.Markdown and returns an HTML fragment (typically an &lt;article class="markdown-body"&gt; wrapper).
    /// When Mermaid is enabled, Mermaid code blocks are annotated with hashes for incremental rendering.
    /// </summary>
    public static string RenderBodyHtml(string markdown, MarkdownRendererOptions? options = null) {
        options ??= new MarkdownRendererOptions();
        var readerOptions = CreateEffectiveReaderOptions(options);
        var htmlOptions = options.HtmlOptions ?? new HtmlOptions { Kind = HtmlKind.Fragment };

        if (options.NormalizeEscapedNewlines && !string.IsNullOrEmpty(markdown)) {
            markdown = markdown.Replace("\\r\\n", "\n").Replace("\\n", "\n");
        }

        markdown ??= string.Empty;
        markdown = PreprocessMarkdown(markdown, options);

        if (options.MaxMarkdownChars.HasValue && options.MaxMarkdownChars.Value >= 0 && markdown.Length > options.MaxMarkdownChars.Value) {
            int max = options.MaxMarkdownChars.Value;
            switch (options.MarkdownOverflowHandling) {
                case OverflowHandling.Throw:
                    throw new ArgumentOutOfRangeException(nameof(markdown), $"Markdown length {markdown.Length} exceeds MaxMarkdownChars {max}.");
                case OverflowHandling.RenderError:
                    return BuildOverflowBodyHtml(htmlOptions, $"Content exceeded the maximum allowed Markdown length ({max} chars).");
                case OverflowHandling.Truncate:
                default:
                    markdown = markdown.Substring(0, max);
                    break;
            }
        }

        if (!string.IsNullOrWhiteSpace(options.BaseHref) && htmlOptions.BaseUri == null) {
            // Best-effort: use BaseHref for origin restrictions (if enabled). If parsing fails or BaseHref isn't absolute,
            // keep BaseUri null and origin restriction will effectively be disabled.
            if (Uri.TryCreate(options.BaseHref!.Trim(), UriKind.Absolute, out var baseUri)) {
                htmlOptions.BaseUri = baseUri;
            }
        }

        var doc = MarkdownReader.Parse(markdown, readerOptions);
        var priorCodeBlockHtmlRenderer = htmlOptions.CodeBlockHtmlRenderer;
        var priorSemanticFencedBlockHtmlRenderer = htmlOptions.SemanticFencedBlockHtmlRenderer;
        htmlOptions.CodeBlockHtmlRenderer = CreateEffectiveCodeBlockHtmlRenderer(options, priorCodeBlockHtmlRenderer);
        htmlOptions.SemanticFencedBlockHtmlRenderer = CreateEffectiveSemanticFencedBlockHtmlRenderer(options, priorSemanticFencedBlockHtmlRenderer);

        string html;
        try {
            html = doc.ToHtmlFragment(htmlOptions) ?? string.Empty;
        } finally {
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

    private static MarkdownCodeBlockHtmlRenderer? CreateEffectiveCodeBlockHtmlRenderer(
        MarkdownRendererOptions options,
        MarkdownCodeBlockHtmlRenderer? priorRenderer) {
        return (block, htmlOptions) => {
            if (priorRenderer != null) {
                var priorHtml = priorRenderer(block, htmlOptions);
                if (priorHtml != null) {
                    return priorHtml;
                }
            }

            return TryRenderCodeBlockOverride(block, options);
        };
    }

    private static MarkdownSemanticFencedBlockHtmlRenderer? CreateEffectiveSemanticFencedBlockHtmlRenderer(
        MarkdownRendererOptions options,
        MarkdownSemanticFencedBlockHtmlRenderer? priorRenderer) {
        return (block, htmlOptions) => {
            if (priorRenderer != null) {
                var priorHtml = priorRenderer(block, htmlOptions);
                if (priorHtml != null) {
                    return priorHtml;
                }
            }

            return TryRenderSemanticFencedBlockOverride(block, options);
        };
    }

    private static string? TryRenderCodeBlockOverride(CodeBlock block, MarkdownRendererOptions options) {
        if (block == null) {
            return null;
        }

        string? replacement = TryRenderCustomFencedCodeBlock(block, options);
        if (replacement == null && options.Mermaid?.Enabled == true && string.Equals(block.Language, "mermaid", StringComparison.OrdinalIgnoreCase)) {
            replacement = BuildMermaidCodeBlockHtml(block.Content, options.Mermaid.EnableHashCaching);
        }

        if (replacement == null && options.Chart?.Enabled == true && string.Equals(block.Language, "chart", StringComparison.OrdinalIgnoreCase)) {
            replacement = BuildChartCodeBlockHtml(block.Content);
        }

        if (replacement == null
            && options.Math?.Enabled == true
            && options.Math.EnableFencedMathBlocks
            && IsMathFenceLanguageAllowed(block.Language, options.Math)) {
            replacement = BuildMathCodeBlockHtml(block.Content);
        }

        if (replacement == null) {
            return null;
        }

        return replacement + BuildCodeBlockCaptionHtml(block.Caption);
    }

    private static string? TryRenderSemanticFencedBlockOverride(SemanticFencedBlock block, MarkdownRendererOptions options) {
        if (block == null) {
            return null;
        }

        string? replacement = TryRenderCustomSemanticFencedBlock(block, options);
        if (replacement == null && options.Mermaid?.Enabled == true && string.Equals(block.SemanticKind, MarkdownSemanticKinds.Mermaid, StringComparison.OrdinalIgnoreCase)) {
            replacement = BuildMermaidCodeBlockHtml(block.Content, options.Mermaid.EnableHashCaching);
        }

        if (replacement == null
            && options.Math?.Enabled == true
            && options.Math.EnableFencedMathBlocks
            && string.Equals(block.SemanticKind, MarkdownSemanticKinds.Math, StringComparison.OrdinalIgnoreCase)) {
            replacement = BuildMathCodeBlockHtml(block.Content);
        }

        if (replacement != null) {
            return replacement + BuildCodeBlockCaptionHtml(block.Caption);
        }

        var codeBlock = new CodeBlock(block.Language, block.Content) {
            Caption = block.Caption
        };
        return TryRenderCodeBlockOverride(codeBlock, options);
    }

    private static string? TryRenderCustomSemanticFencedBlock(SemanticFencedBlock block, MarkdownRendererOptions options) {
        var renderers = options.FencedCodeBlockRenderers;
        if (renderers == null || renderers.Count == 0) {
            return null;
        }

        var exactLanguageMatch = TryRenderMatchingFencedCodeBlockRenderer(
            renderers,
            renderer => RendererHandlesLanguage(renderer, block.Language),
            CreateCodeBlockMatch(block.Language, block.Content),
            options);
        if (exactLanguageMatch != null) {
            return exactLanguageMatch;
        }

        return TryRenderMatchingFencedCodeBlockRenderer(
            renderers,
            renderer => RendererHandlesSemanticKind(renderer, block.SemanticKind),
            CreateCodeBlockMatch(block.Language, block.Content),
            options);
    }

    private static string? TryRenderCustomFencedCodeBlock(CodeBlock block, MarkdownRendererOptions options) {
        var renderers = options.FencedCodeBlockRenderers;
        if (renderers == null || renderers.Count == 0) {
            return null;
        }

        return TryRenderMatchingFencedCodeBlockRenderer(
            renderers,
            renderer => RendererHandlesLanguage(renderer, block.Language),
            CreateCodeBlockMatch(block),
            options);
    }

    private static string? TryRenderMatchingFencedCodeBlockRenderer(
        IReadOnlyList<MarkdownFencedCodeBlockRenderer> renderers,
        Func<MarkdownFencedCodeBlockRenderer, bool> predicate,
        MarkdownFencedCodeBlockMatch match,
        MarkdownRendererOptions options) {
        if (renderers == null || renderers.Count == 0) {
            return null;
        }

        for (int i = renderers.Count - 1; i >= 0; i--) {
            var renderer = renderers[i];
            if (renderer == null || !predicate(renderer)) {
                continue;
            }

            var replacement = renderer.RenderHtml(match, options);
            if (replacement != null) {
                return replacement;
            }
        }

        return null;
    }

    private static bool RendererHandlesLanguage(MarkdownFencedCodeBlockRenderer renderer, string language) {
        var languages = renderer.Languages;
        if (languages == null || languages.Count == 0) {
            return false;
        }

        for (int i = 0; i < languages.Count; i++) {
            var candidate = languages[i];
            if (!string.IsNullOrWhiteSpace(candidate) && string.Equals(candidate, language, StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
        }

        return false;
    }

    private static bool RendererHandlesSemanticKind(MarkdownFencedCodeBlockRenderer renderer, string semanticKind) {
        return !string.IsNullOrWhiteSpace(renderer.SemanticKind)
               && !string.IsNullOrWhiteSpace(semanticKind)
               && string.Equals(renderer.SemanticKind, semanticKind, StringComparison.OrdinalIgnoreCase);
    }

    private static MarkdownFencedCodeBlockMatch CreateCodeBlockMatch(CodeBlock block) {
        return CreateCodeBlockMatch(block.Language, block.Content);
    }

    private static MarkdownFencedCodeBlockMatch CreateCodeBlockMatch(string language, string rawContent) {
        rawContent ??= string.Empty;
        var encodedContent = BuildHtmlEncodedCodeBlockContent(rawContent);
        var originalHtml = BuildDefaultCodeBlockPreHtml(language, rawContent);
        return new MarkdownFencedCodeBlockMatch(language, encodedContent, rawContent, originalHtml);
    }

    private static string BuildHtmlEncodedCodeBlockContent(string rawContent) {
        var encodedContent = System.Net.WebUtility.HtmlEncode(rawContent ?? string.Empty);
        if (encodedContent.Length > 0) {
            encodedContent += "\n";
        }

        return encodedContent;
    }

    private static string BuildDefaultCodeBlockPreHtml(string language, string rawContent) {
        var encodedLanguage = string.IsNullOrEmpty(language)
            ? string.Empty
            : $" class=\"language-{System.Net.WebUtility.HtmlEncode(language)}\"";
        var encodedContent = BuildHtmlEncodedCodeBlockContent(rawContent);
        return $"<pre><code{encodedLanguage}>{encodedContent}</code></pre>";
    }

    private static string BuildCodeBlockCaptionHtml(string? caption) {
        return string.IsNullOrWhiteSpace(caption)
            ? string.Empty
            : $"<div class=\"caption\">{System.Net.WebUtility.HtmlEncode(caption)}</div>";
    }

    private static string BuildMermaidCodeBlockHtml(string rawContent, bool enableHashCaching) {
        var encodedContent = System.Net.WebUtility.HtmlEncode(rawContent ?? string.Empty);
        string hashAttr = string.Empty;
        if (enableHashCaching) {
            string hash = ComputeShortHash(rawContent ?? string.Empty);
            hashAttr = $" data-mermaid-hash=\"{hash}\"";
        }

        return $"<pre class=\"mermaid\"{hashAttr}>{encodedContent}</pre>";
    }

    private static string BuildChartCodeBlockHtml(string rawJson) {
        var payload = MarkdownVisualContract.CreatePayload(rawJson);
        return MarkdownVisualContract.BuildElementHtml(
            "canvas",
            "omd-visual omd-chart",
            MarkdownSemanticKinds.Chart,
            MarkdownSemanticKinds.Chart,
            payload,
            new KeyValuePair<string, string?>("data-chart-hash", payload.Hash),
            new KeyValuePair<string, string?>("data-chart-config-b64", payload.Base64));
    }

    private static string BuildMathCodeBlockHtml(string rawContent) {
        var safe = System.Net.WebUtility.HtmlEncode(rawContent ?? string.Empty);
        return "<div class=\"omd-math\">$$\n" + safe + "\n$$</div>";
    }

    private static string PreprocessMarkdown(string markdown, MarkdownRendererOptions options) {
        var value = markdown ?? string.Empty;
        if (value.Length == 0) {
            return value;
        }

        var normalization = CreateEffectiveInputNormalization(options);
        var preParseNormalization = CreatePreParseNormalizationOptions(normalization);
        if (preParseNormalization != null) {
            value = MarkdownInputNormalizer.Normalize(value, preParseNormalization);
        }

        var pre = options.MarkdownPreProcessors;
        if (pre != null && pre.Count > 0) {
            for (int i = 0; i < pre.Count; i++) {
                var processor = pre[i];
                if (processor == null) continue;
                value = processor(value, options) ?? value ?? string.Empty;
            }
        }

        return value;
    }

    private static MarkdownReaderOptions CreateEffectiveReaderOptions(MarkdownRendererOptions options) {
        var source = options.ReaderOptions ?? new MarkdownReaderOptions();
        var normalization = CreateEffectiveInputNormalization(options);
        var effective = new MarkdownReaderOptions {
            FrontMatter = source.FrontMatter,
            Callouts = source.Callouts,
            Headings = source.Headings,
            FencedCode = source.FencedCode,
            IndentedCodeBlocks = source.IndentedCodeBlocks,
            Images = source.Images,
            UnorderedLists = source.UnorderedLists,
            OrderedLists = source.OrderedLists,
            TaskLists = source.TaskLists,
            Tables = source.Tables,
            DefinitionLists = source.DefinitionLists,
            TocPlaceholders = source.TocPlaceholders,
            Footnotes = source.Footnotes,
            PreferNarrativeSingleLineDefinitions = source.PreferNarrativeSingleLineDefinitions,
            HtmlBlocks = source.HtmlBlocks,
            Paragraphs = source.Paragraphs,
            AutolinkUrls = source.AutolinkUrls,
            AutolinkWwwUrls = source.AutolinkWwwUrls,
            AutolinkWwwScheme = source.AutolinkWwwScheme,
            AutolinkEmails = source.AutolinkEmails,
            BackslashHardBreaks = source.BackslashHardBreaks,
            InlineHtml = source.InlineHtml,
            BaseUri = source.BaseUri,
            DisallowScriptUrls = source.DisallowScriptUrls,
            DisallowFileUrls = source.DisallowFileUrls,
            AllowMailtoUrls = source.AllowMailtoUrls,
            AllowDataUrls = source.AllowDataUrls,
            AllowProtocolRelativeUrls = source.AllowProtocolRelativeUrls,
            RestrictUrlSchemes = source.RestrictUrlSchemes,
            AllowedUrlSchemes = source.AllowedUrlSchemes,
            InputNormalization = CreateInlineNormalizationOptions(normalization)
        };

        AddRendererSemanticFencedBlockExtensions(effective, options);
        CopyBlockParserExtensions(source, effective);
        CopyFencedBlockExtensions(source, effective);
        CopyDocumentTransforms(source, effective);
        return effective;
    }

    private static MarkdownInputNormalizationOptions CreateEffectiveInputNormalization(MarkdownRendererOptions options) {
        var source = options.ReaderOptions?.InputNormalization;
        return new MarkdownInputNormalizationOptions {
            NormalizeSoftWrappedStrongSpans = (source?.NormalizeSoftWrappedStrongSpans == true) || options.NormalizeSoftWrappedStrongSpans,
            NormalizeInlineCodeSpanLineBreaks = (source?.NormalizeInlineCodeSpanLineBreaks == true) || options.NormalizeInlineCodeSpanLineBreaks,
            NormalizeEscapedInlineCodeSpans = (source?.NormalizeEscapedInlineCodeSpans == true) || options.NormalizeEscapedInlineCodeSpans,
            NormalizeTightStrongBoundaries = (source?.NormalizeTightStrongBoundaries == true) || options.NormalizeTightStrongBoundaries,
            NormalizeTightArrowStrongBoundaries = (source?.NormalizeTightArrowStrongBoundaries == true) || options.NormalizeTightArrowStrongBoundaries,
            NormalizeBrokenStrongArrowLabels = (source?.NormalizeBrokenStrongArrowLabels == true) || options.NormalizeBrokenStrongArrowLabels,
            NormalizeWrappedSignalFlowStrongRuns = (source?.NormalizeWrappedSignalFlowStrongRuns == true) || options.NormalizeWrappedSignalFlowStrongRuns,
            NormalizeSignalFlowLabelSpacing = (source?.NormalizeSignalFlowLabelSpacing == true) || options.NormalizeSignalFlowLabelSpacing,
            NormalizeCollapsedMetricChains = (source?.NormalizeCollapsedMetricChains == true) || options.NormalizeCollapsedMetricChains,
            NormalizeHostLabelBulletArtifacts = (source?.NormalizeHostLabelBulletArtifacts == true) || options.NormalizeHostLabelBulletArtifacts,
            NormalizeTightColonSpacing = (source?.NormalizeTightColonSpacing == true) || options.NormalizeTightColonSpacing,
            NormalizeHeadingListBoundaries = (source?.NormalizeHeadingListBoundaries == true) || options.NormalizeHeadingListBoundaries,
            NormalizeCompactStrongLabelListBoundaries = (source?.NormalizeCompactStrongLabelListBoundaries == true) || options.NormalizeCompactStrongLabelListBoundaries,
            NormalizeCompactHeadingBoundaries = (source?.NormalizeCompactHeadingBoundaries == true) || options.NormalizeCompactHeadingBoundaries,
            NormalizeStandaloneHashHeadingSeparators = (source?.NormalizeStandaloneHashHeadingSeparators == true) || options.NormalizeStandaloneHashHeadingSeparators,
            NormalizeBrokenTwoLineStrongLeadIns = (source?.NormalizeBrokenTwoLineStrongLeadIns == true) || options.NormalizeBrokenTwoLineStrongLeadIns,
            NormalizeColonListBoundaries = (source?.NormalizeColonListBoundaries == true) || options.NormalizeColonListBoundaries,
            NormalizeCompactFenceBodyBoundaries = (source?.NormalizeCompactFenceBodyBoundaries == true) || options.NormalizeCompactFenceBodyBoundaries,
            NormalizeLooseStrongDelimiters = (source?.NormalizeLooseStrongDelimiters == true) || options.NormalizeLooseStrongDelimiters,
            NormalizeOrderedListMarkerSpacing = (source?.NormalizeOrderedListMarkerSpacing == true) || options.NormalizeOrderedListMarkerSpacing,
            NormalizeOrderedListParenMarkers = (source?.NormalizeOrderedListParenMarkers == true) || options.NormalizeOrderedListParenMarkers,
            NormalizeOrderedListCaretArtifacts = (source?.NormalizeOrderedListCaretArtifacts == true) || options.NormalizeOrderedListCaretArtifacts,
            NormalizeCollapsedOrderedListBoundaries = (source?.NormalizeCollapsedOrderedListBoundaries == true) || options.NormalizeCollapsedOrderedListBoundaries,
            NormalizeOrderedListStrongDetailClosures = (source?.NormalizeOrderedListStrongDetailClosures == true) || options.NormalizeOrderedListStrongDetailClosures,
            NormalizeTightParentheticalSpacing = (source?.NormalizeTightParentheticalSpacing == true) || options.NormalizeTightParentheticalSpacing,
            NormalizeNestedStrongDelimiters = (source?.NormalizeNestedStrongDelimiters == true) || options.NormalizeNestedStrongDelimiters,
            NormalizeDanglingTrailingStrongListClosers = (source?.NormalizeDanglingTrailingStrongListClosers == true) || options.NormalizeDanglingTrailingStrongListClosers,
            NormalizeMetricValueStrongRuns = (source?.NormalizeMetricValueStrongRuns == true) || options.NormalizeMetricValueStrongRuns
        };
    }

    private static MarkdownInputNormalizationOptions CreateInlineNormalizationOptions(MarkdownInputNormalizationOptions source) {
        return new MarkdownInputNormalizationOptions {
            NormalizeEscapedInlineCodeSpans = source?.NormalizeEscapedInlineCodeSpans ?? false,
            NormalizeTightStrongBoundaries = source?.NormalizeTightStrongBoundaries ?? false,
            NormalizeTightColonSpacing = source?.NormalizeTightColonSpacing ?? false
        };
    }

    private static void CopyFencedBlockExtensions(MarkdownReaderOptions source, MarkdownReaderOptions target) {
        var extensions = source.FencedBlockExtensions;
        if (extensions == null || extensions.Count == 0) {
            return;
        }

        for (int i = 0; i < extensions.Count; i++) {
            var extension = extensions[i];
            if (extension != null) {
                target.FencedBlockExtensions.Add(extension);
            }
        }
    }

    private static void CopyBlockParserExtensions(MarkdownReaderOptions source, MarkdownReaderOptions target) {
        var extensions = source.BlockParserExtensions;
        target.BlockParserExtensions.Clear();
        if (extensions == null || extensions.Count == 0) {
            return;
        }

        for (int i = 0; i < extensions.Count; i++) {
            var extension = extensions[i];
            if (extension != null) {
                target.BlockParserExtensions.Add(extension);
            }
        }
    }

    private static void CopyDocumentTransforms(MarkdownReaderOptions source, MarkdownReaderOptions target) {
        var transforms = source.DocumentTransforms;
        if (transforms == null || transforms.Count == 0) {
            return;
        }

        for (int i = 0; i < transforms.Count; i++) {
            var transform = transforms[i];
            if (transform != null) {
                target.DocumentTransforms.Add(transform);
            }
        }
    }

    private static void AddRendererSemanticFencedBlockExtensions(MarkdownReaderOptions target, MarkdownRendererOptions options) {
        AddSemanticFencedBlockExtension(target, "Built-in Mermaid AST", new[] { MarkdownSemanticKinds.Mermaid }, MarkdownSemanticKinds.Mermaid);

        var mathLanguages = options.Math?.FencedMathLanguages;
        if (mathLanguages != null && mathLanguages.Length > 0) {
            AddSemanticFencedBlockExtension(target, "Built-in Math AST", mathLanguages, MarkdownSemanticKinds.Math);
        }

        var renderers = options.FencedCodeBlockRenderers;
        if (renderers == null || renderers.Count == 0) {
            return;
        }

        for (int i = 0; i < renderers.Count; i++) {
            var renderer = renderers[i];
            if (renderer == null) {
                continue;
            }

            var semanticKind = string.IsNullOrWhiteSpace(renderer.SemanticKind)
                ? renderer.Languages[0]
                : renderer.SemanticKind;
            AddSemanticFencedBlockExtension(target, renderer.Name + " AST", renderer.Languages, semanticKind);
        }
    }

    private static void AddSemanticFencedBlockExtension(
        MarkdownReaderOptions target,
        string name,
        IEnumerable<string> languages,
        string semanticKind) {
        target.FencedBlockExtensions.Add(new MarkdownFencedBlockExtension(
            name,
            languages,
            context => new SemanticFencedBlock(semanticKind, context.Language, context.Content, context.Caption)));
    }

    private static MarkdownInputNormalizationOptions? CreatePreParseNormalizationOptions(MarkdownInputNormalizationOptions source) {
        bool normalizeZeroWidthSpacingArtifacts = source?.NormalizeZeroWidthSpacingArtifacts ?? false;
        bool normalizeEmojiWordJoins = source?.NormalizeEmojiWordJoins ?? false;
        bool normalizeCompactNumberedChoiceBoundaries = source?.NormalizeCompactNumberedChoiceBoundaries ?? false;
        bool normalizeSentenceCollapsedBullets = source?.NormalizeSentenceCollapsedBullets ?? false;
        bool normalizeSoftWrappedStrong = source?.NormalizeSoftWrappedStrongSpans ?? false;
        bool normalizeInlineCodeLineBreaks = source?.NormalizeInlineCodeSpanLineBreaks ?? false;
        bool normalizeLooseStrongDelimiters = source?.NormalizeLooseStrongDelimiters ?? false;
        bool normalizeTightStrongBoundaries = source?.NormalizeTightStrongBoundaries ?? false;
        bool normalizeTightArrowStrongBoundaries = source?.NormalizeTightArrowStrongBoundaries ?? false;
        bool normalizeBrokenStrongArrowLabels = source?.NormalizeBrokenStrongArrowLabels ?? false;
        bool normalizeHostLabelBulletArtifacts = source?.NormalizeHostLabelBulletArtifacts ?? false;
        bool normalizeHeadingListBoundaries = source?.NormalizeHeadingListBoundaries ?? false;
        bool normalizeCompactStrongLabelListBoundaries = source?.NormalizeCompactStrongLabelListBoundaries ?? false;
        bool normalizeCompactHeadingBoundaries = source?.NormalizeCompactHeadingBoundaries ?? false;
        bool normalizeStandaloneHashHeadingSeparators = source?.NormalizeStandaloneHashHeadingSeparators ?? false;
        bool normalizeBrokenTwoLineStrongLeadIns = source?.NormalizeBrokenTwoLineStrongLeadIns ?? false;
        bool normalizeColonListBoundaries = source?.NormalizeColonListBoundaries ?? false;
        bool normalizeCompactFenceBodyBoundaries = source?.NormalizeCompactFenceBodyBoundaries ?? false;
        bool normalizeOrderedListMarkerSpacing = source?.NormalizeOrderedListMarkerSpacing ?? false;
        bool normalizeOrderedListParenMarkers = source?.NormalizeOrderedListParenMarkers ?? false;
        bool normalizeOrderedListCaretArtifacts = source?.NormalizeOrderedListCaretArtifacts ?? false;
        bool normalizeCollapsedOrderedListBoundaries = source?.NormalizeCollapsedOrderedListBoundaries ?? false;
        bool normalizeOrderedListStrongDetailClosures = source?.NormalizeOrderedListStrongDetailClosures ?? false;
        bool normalizeTightParentheticalSpacing = source?.NormalizeTightParentheticalSpacing ?? false;
        bool normalizeNestedStrongDelimiters = source?.NormalizeNestedStrongDelimiters ?? false;

        if (!normalizeZeroWidthSpacingArtifacts
            && !normalizeEmojiWordJoins
            && !normalizeCompactNumberedChoiceBoundaries
            && !normalizeSentenceCollapsedBullets
            && !normalizeSoftWrappedStrong
            && !normalizeInlineCodeLineBreaks
            && !normalizeLooseStrongDelimiters
            && !normalizeTightStrongBoundaries
            && !normalizeTightArrowStrongBoundaries
            && !normalizeBrokenStrongArrowLabels
            && !normalizeHostLabelBulletArtifacts
            && !normalizeHeadingListBoundaries
            && !normalizeCompactStrongLabelListBoundaries
            && !normalizeCompactHeadingBoundaries
            && !normalizeStandaloneHashHeadingSeparators
            && !normalizeBrokenTwoLineStrongLeadIns
            && !normalizeColonListBoundaries
            && !normalizeCompactFenceBodyBoundaries
            && !normalizeOrderedListMarkerSpacing
            && !normalizeOrderedListParenMarkers
            && !normalizeOrderedListCaretArtifacts
            && !normalizeCollapsedOrderedListBoundaries
            && !normalizeOrderedListStrongDetailClosures
            && !normalizeTightParentheticalSpacing
            && !normalizeNestedStrongDelimiters) {
            return null;
        }

        return new MarkdownInputNormalizationOptions {
            NormalizeZeroWidthSpacingArtifacts = normalizeZeroWidthSpacingArtifacts,
            NormalizeEmojiWordJoins = normalizeEmojiWordJoins,
            NormalizeCompactNumberedChoiceBoundaries = normalizeCompactNumberedChoiceBoundaries,
            NormalizeSentenceCollapsedBullets = normalizeSentenceCollapsedBullets,
            NormalizeSoftWrappedStrongSpans = normalizeSoftWrappedStrong,
            NormalizeInlineCodeSpanLineBreaks = normalizeInlineCodeLineBreaks,
            NormalizeLooseStrongDelimiters = normalizeLooseStrongDelimiters,
            NormalizeTightStrongBoundaries = normalizeTightStrongBoundaries,
            NormalizeTightArrowStrongBoundaries = normalizeTightArrowStrongBoundaries,
            NormalizeBrokenStrongArrowLabels = normalizeBrokenStrongArrowLabels,
            NormalizeHostLabelBulletArtifacts = normalizeHostLabelBulletArtifacts,
            NormalizeHeadingListBoundaries = normalizeHeadingListBoundaries,
            NormalizeCompactStrongLabelListBoundaries = normalizeCompactStrongLabelListBoundaries,
            NormalizeCompactHeadingBoundaries = normalizeCompactHeadingBoundaries,
            NormalizeStandaloneHashHeadingSeparators = normalizeStandaloneHashHeadingSeparators,
            NormalizeBrokenTwoLineStrongLeadIns = normalizeBrokenTwoLineStrongLeadIns,
            NormalizeColonListBoundaries = normalizeColonListBoundaries,
            NormalizeCompactFenceBodyBoundaries = normalizeCompactFenceBodyBoundaries,
            NormalizeOrderedListMarkerSpacing = normalizeOrderedListMarkerSpacing,
            NormalizeOrderedListParenMarkers = normalizeOrderedListParenMarkers,
            NormalizeOrderedListCaretArtifacts = normalizeOrderedListCaretArtifacts,
            NormalizeCollapsedOrderedListBoundaries = normalizeCollapsedOrderedListBoundaries,
            NormalizeOrderedListStrongDetailClosures = normalizeOrderedListStrongDetailClosures,
            NormalizeTightParentheticalSpacing = normalizeTightParentheticalSpacing,
            NormalizeNestedStrongDelimiters = normalizeNestedStrongDelimiters
        };
    }

    private static string BuildOverflowBodyHtml(HtmlOptions htmlOptions, string message) {
        string msg = System.Net.WebUtility.HtmlEncode(message ?? "Content too large.");
        string inner = $"<blockquote class=\"callout warning\" data-omd=\"overflow\"><p>{msg}</p></blockquote>";

        var bodyClass = htmlOptions.BodyClass;
        if (bodyClass != null) {
            bodyClass = bodyClass.Trim();
            if (bodyClass.Length > 0) {
                string cls = System.Net.WebUtility.HtmlEncode(bodyClass);
                return $"<article class=\"{cls}\">{inner}</article>";
            }
        }

        return $"<div data-omd=\"overflow\">{inner}</div>";
    }

    /// <summary>
    /// Builds a self-contained HTML document that preloads CSS and scripts once (Prism/Mermaid),
    /// and exposes a global <c>updateContent(newBodyHtml)</c> function for incremental updates.
    /// </summary>
    public static string BuildShellHtml(string? title = null, MarkdownRendererOptions? options = null) {
        options ??= new MarkdownRendererOptions();
        var htmlOptions = options.HtmlOptions ?? new HtmlOptions { Kind = HtmlKind.Fragment };

        // Build head assets (CSS + optional Prism assets) from OfficeIMO.Markdown.
        // This intentionally uses an empty doc; content is pushed later via updateContent(...).
        var empty = MarkdownDoc.Create();
        var parts = empty.ToHtmlParts(htmlOptions);

        var sb = new StringBuilder(16 * 1024);
        sb.Append("<!DOCTYPE html><html lang=\"en\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">");
        if (!string.IsNullOrWhiteSpace(options.ContentSecurityPolicy)) {
            sb.Append("<meta http-equiv=\"Content-Security-Policy\" content=\"")
              .Append(System.Net.WebUtility.HtmlEncode(options.ContentSecurityPolicy!.Trim()))
              .Append("\">");
        }
        sb.Append("<title>").Append(System.Net.WebUtility.HtmlEncode(title ?? "Markdown")).Append("</title>");
        if (!string.IsNullOrEmpty(parts.Css)) sb.Append("<style>\n").Append(parts.Css).Append("\n</style>");
        if (!string.IsNullOrEmpty(parts.Head)) sb.Append(parts.Head);
        if (!string.IsNullOrWhiteSpace(options.ShellCss)) {
            sb.Append("<style data-omd=\"shell\">")
              .Append("\n")
              .Append(options.ShellCss)
              .Append("\n</style>");
        }

        var assetMode = htmlOptions.AssetMode;

        if (options.Math?.Enabled == true) {
            sb.Append(BuildMathBootstrap(options.Math, assetMode));
        }

        if (options.Mermaid?.Enabled == true) {
            sb.Append(BuildMermaidBootstrap(options.Mermaid, assetMode));
        }

        if (options.Chart?.Enabled == true) {
            sb.Append(BuildChartBootstrap(options.Chart, assetMode));
        }

        AppendCustomShellHeadHtml(sb, options, assetMode);

        sb.Append("</head><body>");
        sb.Append("<div id=\"omdRoot\"></div>");
        sb.Append("<script>\n").Append(BuildIncrementalUpdateScript(options)).Append("\n</script>");
        sb.Append("</body></html>");
        return sb.ToString();
    }

    /// <summary>
    /// Returns a JavaScript snippet that calls <c>updateContent(...)</c> with a properly escaped string literal.
    /// </summary>
    public static string BuildUpdateScript(string bodyHtml) {
        return "updateContent(" + JavaScriptString.SingleQuoted(bodyHtml ?? string.Empty) + ");";
    }

    /// <summary>
    /// Convenience helper for hosts: renders Markdown to an HTML fragment and returns the JavaScript snippet
    /// that updates the shell (calls <c>updateContent(...)</c>).
    /// </summary>
    public static string RenderUpdateScript(string markdown, MarkdownRendererOptions? options = null) {
        var bodyHtml = RenderBodyHtml(markdown ?? string.Empty, options);
        return BuildUpdateScript(bodyHtml);
    }

    /// <summary>
    /// Wraps an existing HTML fragment in a chat "bubble" container (optional).
    /// This is purely a formatting helper: it does not change Markdown parsing rules.
    /// </summary>
    public static string WrapAsChatBubble(string bodyHtml, ChatMessageRole role = ChatMessageRole.Assistant) {
        string roleClass = role switch {
            ChatMessageRole.User => "omd-role-user",
            ChatMessageRole.System => "omd-role-system",
            _ => "omd-role-assistant"
        };

        // bodyHtml is expected to be the output of RenderBodyHtml (typically an <article class="markdown-body"> wrapper).
        // Keep it as-is and add a single outer container so host UIs don't have to author HTML around each message.
        return $"<div class=\"omd-chat-row {roleClass}\"><div class=\"omd-chat-bubble\">{bodyHtml ?? string.Empty}</div></div>";
    }

    /// <summary>
    /// Convenience helper: renders Markdown then wraps the result in a chat bubble.
    /// </summary>
    public static string RenderChatBubbleBodyHtml(string markdown, ChatMessageRole role = ChatMessageRole.Assistant, MarkdownRendererOptions? options = null) {
        var bodyHtml = RenderBodyHtml(markdown ?? string.Empty, options);
        return WrapAsChatBubble(bodyHtml, role);
    }

    /// <summary>
    /// Convenience helper: renders Markdown as a chat bubble and returns an update script snippet.
    /// </summary>
    public static string RenderChatBubbleUpdateScript(string markdown, ChatMessageRole role = ChatMessageRole.Assistant, MarkdownRendererOptions? options = null) {
        return BuildUpdateScript(RenderChatBubbleBodyHtml(markdown, role, options));
    }

    private static bool IsMathFenceLanguageAllowed(string lang, MathOptions mathOptions) {
        if (string.IsNullOrWhiteSpace(lang)) return false;
        if (mathOptions == null) return false;
        var allowed = mathOptions.FencedMathLanguages;
        if (allowed == null || allowed.Length == 0) return true; // treat as enabled for defaults

        for (int i = 0; i < allowed.Length; i++) {
            var a = (allowed[i] ?? string.Empty).Trim();
            if (a.Length == 0) continue;
            if (string.Equals(a, lang, StringComparison.OrdinalIgnoreCase)) return true;
        }
        return false;
    }

    internal static string ComputeShortHash(string input) {
        var data = Encoding.UTF8.GetBytes(input ?? string.Empty);
        byte[] hash;
#if NET8_0_OR_GREATER
        hash = SHA256.HashData(data);
#else
        using (var sha = SHA256.Create()) {
            hash = sha.ComputeHash(data);
        }
#endif
        // Use first 8 bytes as hex = 16 chars, plenty for DOM-diff keys.
        return ToHex(hash, 8);
    }

    private static string ToHex(byte[] bytes, int take) {
        if (bytes == null || bytes.Length == 0) return string.Empty;
        int len = Math.Min(take, bytes.Length);
        var sb = new StringBuilder(len * 2);
        for (int i = 0; i < len; i++) {
            sb.Append(bytes[i].ToString("x2"));
        }
        return sb.ToString();
    }

    private static string BuildMermaidBootstrap(MermaidOptions o, AssetMode assetMode) {
        // Mermaid bootstrap:
        // - Online: ESM import (default)
        // - Offline: non-module script (easier to bundle/host locally)
        string url = (o?.EsmModuleUrl ?? string.Empty).Trim();
        string scriptUrl = (o?.ScriptUrl ?? string.Empty).Trim();
        string light = (o?.LightTheme ?? "default").Trim();
        string dark = (o?.DarkTheme ?? "dark").Trim();
        if (string.IsNullOrEmpty(url) && string.IsNullOrEmpty(scriptUrl)) return string.Empty;

        // Prevent closing the <script> tag if a caller passes a hostile value.
        url = ReplaceScriptCloseSequence(url);
        scriptUrl = ReplaceScriptCloseSequence(scriptUrl);
        light = ReplaceScriptCloseSequence(light);
        dark = ReplaceScriptCloseSequence(dark);

        if (assetMode == AssetMode.Offline && !string.IsNullOrEmpty(scriptUrl)) {
            string src = BuildBundledScriptSrc(scriptUrl, mime: "application/javascript");
            if (string.IsNullOrEmpty(src)) src = scriptUrl;
            src = System.Net.WebUtility.HtmlEncode(src);
            return $@"
<script src=""{src}""></script>
<script>
// Initialize Mermaid once after load (non-module path).
(function(){{
  try {{
    if (!window.mermaid && typeof mermaid !== 'undefined') window.mermaid = mermaid;
    if (window.mermaid && typeof window.mermaid.initialize === 'function') {{
      window.mermaid.initialize({{ startOnLoad: false, theme: window.matchMedia('(prefers-color-scheme: dark)').matches ? {JavaScriptString.SingleQuoted(dark)} : {JavaScriptString.SingleQuoted(light)} }});
    }}
  }} catch(e) {{ }}
}})();
</script>";
        }

        // Default (online): ESM import.
        if (string.IsNullOrEmpty(url)) return string.Empty;
        return $@"
<script type=""module"">
import mermaid from {JavaScriptString.SingleQuoted(url)};
window.mermaid = mermaid;
mermaid.initialize({{ startOnLoad: false, theme: window.matchMedia('(prefers-color-scheme: dark)').matches ? {JavaScriptString.SingleQuoted(dark)} : {JavaScriptString.SingleQuoted(light)} }});
</script>";
    }

    private static string ReplaceScriptCloseSequence(string value) {
        if (string.IsNullOrEmpty(value)) return string.Empty;
        // Avoid embedding a literal "</script" inside script contents.
        return value.Replace("</", "<\\/");
    }

    private static string BuildChartBootstrap(ChartOptions o, AssetMode assetMode) {
        string url = (o?.ScriptUrl ?? string.Empty).Trim();
        if (string.IsNullOrEmpty(url)) return string.Empty;

        string src = assetMode == AssetMode.Offline ? BuildBundledScriptSrc(url, mime: "application/javascript") : string.Empty;
        if (string.IsNullOrEmpty(src)) src = url;
        src = System.Net.WebUtility.HtmlEncode(src);
        return $"\n<script defer src=\"{src}\"></script>\n";
    }

    private static string BuildMathBootstrap(MathOptions o, AssetMode assetMode) {
        string css = (o?.CssUrl ?? string.Empty).Trim();
        string js = (o?.ScriptUrl ?? string.Empty).Trim();
        string ar = (o?.AutoRenderScriptUrl ?? string.Empty).Trim();
        if (string.IsNullOrEmpty(css) || string.IsNullOrEmpty(js) || string.IsNullOrEmpty(ar)) return string.Empty;

        // KaTeX should be ready before we render content via updateContent(...). Use defer so it doesn't block HTML parse,
        // and call renderMathInElement from updateContent after DOM updates.
        string cssHref = assetMode == AssetMode.Offline ? BuildBundledCssHref(css) : string.Empty;
        if (string.IsNullOrEmpty(cssHref)) cssHref = css;
        cssHref = System.Net.WebUtility.HtmlEncode(cssHref);

        string jsSrc = assetMode == AssetMode.Offline ? BuildBundledScriptSrc(js, mime: "application/javascript") : string.Empty;
        if (string.IsNullOrEmpty(jsSrc)) jsSrc = js;
        jsSrc = System.Net.WebUtility.HtmlEncode(jsSrc);

        string arSrc = assetMode == AssetMode.Offline ? BuildBundledScriptSrc(ar, mime: "application/javascript") : string.Empty;
        if (string.IsNullOrEmpty(arSrc)) arSrc = ar;
        arSrc = System.Net.WebUtility.HtmlEncode(arSrc);

        return $"\n<link rel=\"stylesheet\" href=\"{cssHref}\">\n<script defer src=\"{jsSrc}\"></script>\n<script defer src=\"{arSrc}\"></script>\n";
    }

    private static void AppendCustomShellHeadHtml(StringBuilder sb, MarkdownRendererOptions options, AssetMode assetMode) {
        var renderers = options.FencedCodeBlockRenderers;
        if (renderers == null || renderers.Count == 0) {
            return;
        }

        for (int i = 0; i < renderers.Count; i++) {
            var renderer = renderers[i];
            if (renderer?.BuildShellHeadHtml == null) {
                continue;
            }

            var fragment = renderer.BuildShellHeadHtml(options, assetMode);
            if (!string.IsNullOrWhiteSpace(fragment)) {
                sb.Append(fragment);
            }
        }
    }

    private static void AppendCustomUpdateScripts(StringBuilder sb, MarkdownRendererOptions options, bool beforeReplace) {
        var renderers = options.FencedCodeBlockRenderers;
        if (renderers == null || renderers.Count == 0) {
            return;
        }

        for (int i = 0; i < renderers.Count; i++) {
            var renderer = renderers[i];
            if (renderer == null) {
                continue;
            }

            var builder = beforeReplace
                ? renderer.BuildBeforeContentReplaceScript
                : renderer.BuildAfterContentReplaceScript;
            if (builder == null) {
                continue;
            }

            var fragment = builder(options);
            if (string.IsNullOrWhiteSpace(fragment)) {
                continue;
            }

            sb.Append('\n')
              .Append(ReplaceScriptCloseSequence(fragment ?? string.Empty))
              .Append('\n');
        }
    }

    internal static string BuildBundledScriptSrc(string hrefOrPath, string mime) {
        // Only used by shell building logic. This should never throw.
        try {
            var text = TryLoadTextAsset(hrefOrPath);
            if (string.IsNullOrEmpty(text)) return string.Empty;
            var bytes = Encoding.UTF8.GetBytes(text);
            var b64 = Convert.ToBase64String(bytes);
            return $"data:{mime};base64,{b64}";
        } catch { return string.Empty; }
    }

    internal static string BuildBundledCssHref(string hrefOrPath) {
        try {
            var text = TryLoadTextAsset(hrefOrPath);
            if (string.IsNullOrEmpty(text)) return string.Empty;
            var bytes = Encoding.UTF8.GetBytes(text);
            var b64 = Convert.ToBase64String(bytes);
            return $"data:text/css;base64,{b64}";
        } catch { return string.Empty; }
    }

    private static string TryLoadTextAsset(string hrefOrPath) {
        try {
            if (string.IsNullOrWhiteSpace(hrefOrPath)) return string.Empty;
            string v = hrefOrPath.Trim();

            if (Uri.TryCreate(v, UriKind.Absolute, out var uri)) {
                if (uri.IsFile) {
                    string path = uri.LocalPath;
                    return TryReadAllTextBounded(path);
                }

                if (string.Equals(uri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)) {
                    return TryDownloadTextBounded(uri);
                }

                // Unknown scheme (e.g., custom WebView virtual hosts) cannot be resolved by this process.
                return string.Empty;
            }

            // Treat as local path (absolute or relative).
            return TryReadAllTextBounded(v);
        } catch {
            return string.Empty;
        }
    }

    private static string TryReadAllTextBounded(string path) {
        try {
            if (string.IsNullOrWhiteSpace(path)) return string.Empty;
            if (!System.IO.File.Exists(path)) return string.Empty;
            const long MaxBytes = 10_000_000; // 10MB guardrail
            var fi = new System.IO.FileInfo(path);
            if (fi.Length > MaxBytes) return string.Empty;
            return System.IO.File.ReadAllText(path, Encoding.UTF8);
        } catch {
            return string.Empty;
        }
    }

    private static string TryDownloadTextBounded(Uri uri) {
        try {
            if (uri == null) return string.Empty;

            const long MaxBytes = 10_000_000; // 10MB guardrail
            using var cts = new System.Threading.CancellationTokenSource(TimeSpan.FromSeconds(8));
            using var client = new System.Net.Http.HttpClient();

            using var resp = client.GetAsync(uri, System.Net.Http.HttpCompletionOption.ResponseHeadersRead, cts.Token).GetAwaiter().GetResult();
            if (!resp.IsSuccessStatusCode) return string.Empty;

            var len = resp.Content.Headers.ContentLength;
            if (len.HasValue && len.Value > MaxBytes) return string.Empty;

            using var stream = resp.Content.ReadAsStreamAsync().GetAwaiter().GetResult();
            using var mem = new System.IO.MemoryStream(len.HasValue ? (int)Math.Min(len.Value, MaxBytes) : 64 * 1024);
            var buffer = new byte[81920];
            long total = 0;
            while (true) {
                int read = stream.Read(buffer, 0, buffer.Length);
                if (read <= 0) break;
                total += read;
                if (total > MaxBytes) return string.Empty;
                mem.Write(buffer, 0, read);
            }
            return Encoding.UTF8.GetString(mem.ToArray());
        } catch {
            return string.Empty;
        }
    }

    private static string BuildIncrementalUpdateScript(MarkdownRendererOptions options) {
        bool mermaid = options.Mermaid?.Enabled == true;
        bool chart = options.Chart?.Enabled == true;
        var mathOptions = options.Math;
        bool codeCopy = options.EnableCodeCopyButtons;
        bool tableCopy = options.EnableTableCopyButtons;

        // Notes:
        // - We keep <base> in <head> so relative links/images resolve.
        // - We preserve already-rendered Mermaid SVGs by comparing data-mermaid-hash attributes.
        // - We re-run Prism highlighting after updates (if Prism is present).
        var sb = new StringBuilder(8 * 1024);
        sb.Append("""
async function updateContent(newBodyHtml) {
  const root = document.getElementById('omdRoot') || document.body;
  // Extract <base href="..."> from payload and place it in <head>.
  try {
    const baseMatch = newBodyHtml.match(/<base\s+href="([^"]*)"[^>]*>/i);
    if (baseMatch) {
      let baseEl = document.getElementById('omdBase');
      if (!baseEl) {
        baseEl = document.createElement('base');
        baseEl.id = 'omdBase';
        document.head.appendChild(baseEl);
      }
      baseEl.href = baseMatch[1];
      newBodyHtml = newBodyHtml.replace(baseMatch[0], '');
    } else {
      const baseEl = document.getElementById('omdBase');
      if (baseEl) baseEl.href = 'about:blank';
    }
  } catch(e) { /* best-effort */ }
""");

        if (chart) {
            sb.Append("""
  // Destroy existing Chart.js instances before replacing DOM to avoid leaks.
  try {
    if (window.Chart && typeof Chart.getChart === 'function') {
      root.querySelectorAll('canvas.omd-chart').forEach(c => {
        try { const inst = Chart.getChart(c); if (inst) inst.destroy(); } catch(e) { /* ignore */ }
      });
    }
  } catch(e) { /* ignore */ }
""");
        }

        if (mermaid) {
            sb.Append("""
  // Cache existing Mermaid SVGs keyed by data-mermaid-hash.
  const existingSvgs = new Map();
  root.querySelectorAll('[data-mermaid-hash]').forEach(el => {
    const hash = el.getAttribute('data-mermaid-hash');
    const svg = el.querySelector('svg') || (el.nextElementSibling && el.nextElementSibling.tagName === 'svg' ? el.nextElementSibling : null);
    if (hash && svg) existingSvgs.set(hash, svg.cloneNode(true));
  });
""");
        }

        AppendCustomUpdateScripts(sb, options, beforeReplace: true);

        sb.Append("""
  // Replace rendered contents.
  root.innerHTML = newBodyHtml;
""");

        if (codeCopy || tableCopy) {
            sb.Append("""

  // Copy helpers (optional)
  function omdCopyText(text) {
    const s = String(text ?? '');
    try {
      const wv = window.chrome && window.chrome.webview;
      if (wv && typeof wv.postMessage === 'function') {
        // Host can optionally handle this message and place text on clipboard.
        wv.postMessage({ type: 'omd.copy', text: s });
      }
    } catch(_) { /* ignore */ }

    try {
      if (navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
        return navigator.clipboard.writeText(s);
      }
    } catch(_) { /* ignore */ }

    try {
      const ta = document.createElement('textarea');
      ta.value = s;
      ta.setAttribute('readonly', 'readonly');
      ta.style.position = 'fixed';
      ta.style.left = '-9999px';
      document.body.appendChild(ta);
      ta.select();
      try { document.execCommand('copy'); } catch(_) { /* ignore */ }
      document.body.removeChild(ta);
    } catch(_) { /* ignore */ }

    return Promise.resolve();
  }

  function omdFlash(btn, label) {
    try {
      const orig = btn.textContent;
      btn.textContent = label;
      btn.setAttribute('data-omd-flash', '1');
      setTimeout(() => { try { btn.textContent = orig; btn.removeAttribute('data-omd-flash'); } catch(_){} }, 900);
    } catch(_) {}
  }
""");
        }

        if (codeCopy) {
            sb.Append("""

  function omdSetupCodeCopyButtons(rootEl) {
    try {
      rootEl.querySelectorAll('pre > code').forEach(code => {
        const pre = code.parentElement;
        if (!pre || pre.getAttribute('data-omd-code-inited') === '1') return;
        pre.setAttribute('data-omd-code-inited', '1');
        pre.classList && pre.classList.add('omd-has-actions');

        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'omd-copy-btn omd-copy-code';
        btn.textContent = 'Copy';
        btn.addEventListener('click', ev => {
          try { ev.preventDefault(); ev.stopPropagation(); } catch(_) {}
          omdCopyText(code.textContent || '');
          omdFlash(btn, 'Copied');
        });

        // Put the button as the first child so it stays visible even if Prism modifies <code>.
        pre.insertBefore(btn, pre.firstChild);
      });
    } catch(_) { /* ignore */ }
  }
""");
        }

        if (tableCopy) {
            sb.Append("""

  function omdCellText(cell) {
    const t = (cell && (cell.innerText || cell.textContent)) ? String(cell.innerText || cell.textContent) : '';
    return t.replace(/\r?\n/g, ' ').trim();
  }

  function omdTableToTsv(table) {
    const rows = [];
    const trs = table.querySelectorAll('tr');
    trs.forEach(tr => {
      const cells = tr.querySelectorAll('th,td');
      if (!cells || cells.length === 0) return;
      const vals = [];
      cells.forEach(c => vals.push(omdCellText(c)));
      rows.push(vals.join('\\t'));
    });
    return rows.join('\\n');
  }

  function omdCsvEscape(value) {
    const s = String(value ?? '');
    if (s.indexOf('\"') >= 0 || s.indexOf(',') >= 0 || s.indexOf('\\n') >= 0 || s.indexOf('\\r') >= 0) {
      return '\"' + s.replace(/\"/g, '\"\"') + '\"';
    }
    return s;
  }

  function omdTableToCsv(table) {
    const rows = [];
    const trs = table.querySelectorAll('tr');
    trs.forEach(tr => {
      const cells = tr.querySelectorAll('th,td');
      if (!cells || cells.length === 0) return;
      const vals = [];
      cells.forEach(c => vals.push(omdCsvEscape(omdCellText(c))));
      rows.push(vals.join(','));
    });
    return rows.join('\\n');
  }

  function omdSetupTableCopyButtons(rootEl) {
    try {
      rootEl.querySelectorAll('table').forEach(table => {
        if (table.getAttribute('data-omd-table-inited') === '1') return;
        table.setAttribute('data-omd-table-inited', '1');

        const actions = document.createElement('div');
        actions.className = 'omd-table-actions';

        const b1 = document.createElement('button');
        b1.type = 'button';
        b1.className = 'omd-copy-btn omd-copy-tsv';
        b1.textContent = 'Copy TSV';
        b1.addEventListener('click', ev => {
          try { ev.preventDefault(); ev.stopPropagation(); } catch(_) {}
          omdCopyText(omdTableToTsv(table));
          omdFlash(b1, 'Copied');
        });

        const b2 = document.createElement('button');
        b2.type = 'button';
        b2.className = 'omd-copy-btn omd-copy-csv';
        b2.textContent = 'Copy CSV';
        b2.addEventListener('click', ev => {
          try { ev.preventDefault(); ev.stopPropagation(); } catch(_) {}
          omdCopyText(omdTableToCsv(table));
          omdFlash(b2, 'Copied');
        });

        actions.appendChild(b1);
        actions.appendChild(b2);

        table.parentElement && table.parentElement.insertBefore(actions, table);
      });
    } catch(_) { /* ignore */ }
  }
""");
        }

        if (mermaid) {
            sb.Append("""
  // Restore cached Mermaid SVGs for unchanged diagrams.
  root.querySelectorAll('[data-mermaid-hash]').forEach(el => {
    const hash = el.getAttribute('data-mermaid-hash');
    if (existingSvgs.has(hash)) {
      const cachedSvg = existingSvgs.get(hash);
      el.innerHTML = '';
      el.appendChild(cachedSvg);
      el.setAttribute('data-mermaid-rendered', 'true');
    }
  });

  // Render only new/changed Mermaid blocks.
  const unrendered = root.querySelectorAll('[data-mermaid-hash]:not([data-mermaid-rendered])');
  if (unrendered.length > 0 && window.mermaid) {
    try { await window.mermaid.run({ nodes: unrendered }); } catch(e) { console.warn('Mermaid render error:', e); }
  }
  // Render plain Mermaid blocks (language-mermaid) when hashes are not present.
  const plainMermaid = root.querySelectorAll('pre > code.language-mermaid:not([data-mermaid-rendered]), .mermaid:not([data-mermaid-rendered]):not(svg)');
  if (plainMermaid.length > 0 && window.mermaid) {
    try { await window.mermaid.run({ nodes: plainMermaid }); } catch(e) { console.warn('Mermaid render error:', e); }
  }
""");
        }

        if (chart) {
            sb.Append("""
  // Chart.js rendering (optional).
  try {
    function b64ToUtf8(b64) {
      try {
        const bytes = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
        if (window.TextDecoder) return new TextDecoder('utf-8').decode(bytes);
        // Fallback for older engines.
        return decodeURIComponent(escape(String.fromCharCode.apply(null, Array.from(bytes))));
      } catch(e) { return ''; }
    }
    if (window.Chart) {
      root.querySelectorAll('canvas.omd-chart:not([data-chart-rendered])').forEach(c => {
        const b64 = c.getAttribute('data-chart-config-b64');
        if (!b64) return;
        const jsonText = b64ToUtf8(b64);
        if (!jsonText) return;
        let cfg = null;
        try { cfg = JSON.parse(jsonText); } catch(e) { console.warn('Chart config JSON parse error:', e); return; }
        try {
          const ctx = c.getContext && c.getContext('2d');
          if (!ctx) return;
          new Chart(ctx, cfg);
          c.setAttribute('data-omd-visual-rendered', 'true');
          c.setAttribute('data-chart-rendered', 'true');
        } catch(e) { console.warn('Chart render error:', e); }
      });
    }
  } catch(e) { /* ignore */ }
""");
        }

        AppendCustomUpdateScripts(sb, options, beforeReplace: false);

        if (codeCopy || tableCopy) {
            sb.Append("""

  // Add optional copy buttons after updates (best-effort).
  try {
""");
            if (codeCopy) sb.Append("    omdSetupCodeCopyButtons(root);\n");
            if (tableCopy) sb.Append("    omdSetupTableCopyButtons(root);\n");
            sb.Append("""
  } catch(_) { /* ignore */ }
""");
        }

        sb.Append("""
  // Prism highlighting (optional).
  try {
    if (window.Prism) {
      if (typeof Prism.highlightAllUnder === 'function') Prism.highlightAllUnder(root);
      else if (typeof Prism.highlightAll === 'function') Prism.highlightAll();
    }
  } catch(e) { /* ignore */ }
""");

        if (mathOptions != null && mathOptions.Enabled) {
            sb.Append("""

  // KaTeX auto-render (optional).
  try {
    if (window.renderMathInElement) {
      const delimiters = [];
""");
            if (mathOptions.EnableDollarDisplay) sb.Append("      delimiters.push({ left: \"$$\", right: \"$$\", display: true });\n");
            if (mathOptions.EnableDollarInline) sb.Append("      delimiters.push({ left: \"$\", right: \"$\", display: false });\n");
            if (mathOptions.EnableBracketDisplay) sb.Append("      delimiters.push({ left: \"\\\\[\", right: \"\\\\]\", display: true });\n");
            if (mathOptions.EnableParenInline) sb.Append("      delimiters.push({ left: \"\\\\(\", right: \"\\\\)\", display: false });\n");
            sb.Append("""
      if (delimiters.length > 0) {
        window.renderMathInElement(root, {
          delimiters: delimiters,
          throwOnError: false,
          strict: 'ignore',
          ignoredTags: ['script', 'noscript', 'style', 'textarea', 'pre', 'code']
        });
      }
    }
  } catch(e) { /* ignore */ }
""");
        }

        sb.Append("""
}

// Optional WebView2 integration: allow hosts to push updates without ExecuteScriptAsync.
// - PostWebMessageAsString(bodyHtml)  => e.data is a string
// - PostWebMessageAsJson({ bodyHtml }) => e.data is an object
(function(){
  try {
    const wv = window.chrome && window.chrome.webview;
    if (!wv || typeof wv.addEventListener !== 'function') return;
    wv.addEventListener('message', e => {
      try {
        const d = e && e.data;
        if (d && typeof d === 'object' && d.type === 'omd.update' && typeof d.bodyHtml === 'string') { updateContent(d.bodyHtml); return; }
        if (typeof d === 'string') { updateContent(d); return; }
        if (d && typeof d === 'object' && typeof d.bodyHtml === 'string') { updateContent(d.bodyHtml); return; }
      } catch(_) { /* ignore */ }
    });
  } catch(_) { /* ignore */ }
})();
""");
        return sb.ToString();
    }
}
