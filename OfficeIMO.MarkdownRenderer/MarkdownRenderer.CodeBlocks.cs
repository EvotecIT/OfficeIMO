using OfficeIMO.Markdown;

namespace OfficeIMO.MarkdownRenderer;

public static partial class MarkdownRenderer {
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
            replacement = BuildChartCodeBlockHtml(block.Content, block.FenceInfo);
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

        var codeBlock = new CodeBlock(block.InfoString, block.Content) {
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
            CreateCodeBlockMatch(block.InfoString, block.Content),
            options);
        if (exactLanguageMatch != null) {
            return exactLanguageMatch;
        }

        return TryRenderMatchingFencedCodeBlockRenderer(
            renderers,
            renderer => RendererHandlesSemanticKind(renderer, block.SemanticKind),
            CreateCodeBlockMatch(block.InfoString, block.Content),
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
        return CreateCodeBlockMatch(block.InfoString, block.Content);
    }

    private static MarkdownFencedCodeBlockMatch CreateCodeBlockMatch(string infoString, string rawContent) {
        rawContent ??= string.Empty;
        var fenceInfo = MarkdownCodeFenceInfo.Parse(infoString);
        var encodedContent = BuildHtmlEncodedCodeBlockContent(rawContent);
        var originalHtml = BuildDefaultCodeBlockPreHtml(fenceInfo.Language, rawContent);
        return new MarkdownFencedCodeBlockMatch(fenceInfo.InfoString, encodedContent, rawContent, originalHtml);
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

    private static string BuildChartCodeBlockHtml(string rawJson, MarkdownCodeFenceInfo? fenceInfo) {
        var payload = MarkdownVisualContract.CreatePayload(rawJson);
        return MarkdownVisualContract.BuildElementHtml(
            "canvas",
            "omd-visual omd-chart",
            MarkdownSemanticKinds.Chart,
            MarkdownSemanticKinds.Chart,
            payload,
            fenceInfo,
            new KeyValuePair<string, string?>(MarkdownVisualElementContract.AttributeVisualTitle, fenceInfo?.Title),
            new KeyValuePair<string, string?>("data-chart-hash", payload.Hash),
            new KeyValuePair<string, string?>("data-chart-config-b64", payload.Base64));
    }

    private static string BuildMathCodeBlockHtml(string rawContent) {
        var safe = System.Net.WebUtility.HtmlEncode(rawContent ?? string.Empty);
        return "<div class=\"omd-math\">$$\n" + safe + "\n$$</div>";
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
}
