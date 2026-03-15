namespace OfficeIMO.Markdown;

/// <summary>
/// Upgrades known visual fenced code blocks and legacy JSON visual payloads into first-class semantic fenced blocks.
/// </summary>
public sealed class MarkdownJsonVisualCodeBlockTransform : IMarkdownDocumentTransform {
    /// <summary>
    /// Creates a visual-code-block transform.
    /// </summary>
    /// <param name="languageMode">Fence-language strategy for upgraded JSON payloads.</param>
    public MarkdownJsonVisualCodeBlockTransform(
        MarkdownVisualFenceLanguageMode languageMode = MarkdownVisualFenceLanguageMode.PreserveOriginal) {
        LanguageMode = languageMode;
    }

    /// <summary>
    /// Fence-language strategy for upgraded JSON payloads.
    /// </summary>
    public MarkdownVisualFenceLanguageMode LanguageMode { get; }

    /// <inheritdoc />
    public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        if (context == null) {
            throw new ArgumentNullException(nameof(context));
        }

        MarkdownDocumentBlockRewriter.RewriteDocument(document, RewriteBlock);
        return document;
    }

    private IMarkdownBlock RewriteBlock(IMarkdownBlock block) {
        if (block is not CodeBlock codeBlock) {
            return block;
        }

        if (!TryResolveVisualSemantic(codeBlock, out var semanticKind, out var language)) {
            return block;
        }

        return new SemanticFencedBlock(semanticKind, language, codeBlock.Content, codeBlock.Caption);
    }

    private bool TryResolveVisualSemantic(CodeBlock block, out string semanticKind, out string language) {
        semanticKind = string.Empty;
        language = block.Language ?? string.Empty;

        if (TryResolveKnownFenceLanguage(language, out semanticKind)) {
            return true;
        }

        if (!MarkdownJsonVisualPayloadDetector.TryDetectSemanticKind(block.Content, out semanticKind)) {
            return false;
        }

        language = ResolveUpgradedLanguage(semanticKind, language);
        return true;
    }

    private static bool TryResolveKnownFenceLanguage(string language, out string semanticKind) {
        semanticKind = string.Empty;
        if (string.IsNullOrWhiteSpace(language)) {
            return false;
        }

        var normalized = language.Trim();
        if (normalized.Equals("chart", StringComparison.OrdinalIgnoreCase)
            || normalized.Equals("ix-chart", StringComparison.OrdinalIgnoreCase)) {
            semanticKind = MarkdownSemanticKinds.Chart;
            return true;
        }

        if (normalized.Equals("network", StringComparison.OrdinalIgnoreCase)
            || normalized.Equals("visnetwork", StringComparison.OrdinalIgnoreCase)
            || normalized.Equals("ix-network", StringComparison.OrdinalIgnoreCase)) {
            semanticKind = MarkdownSemanticKinds.Network;
            return true;
        }

        if (normalized.Equals("dataview", StringComparison.OrdinalIgnoreCase)
            || normalized.Equals("ix-dataview", StringComparison.OrdinalIgnoreCase)) {
            semanticKind = MarkdownSemanticKinds.DataView;
            return true;
        }

        return false;
    }

    private string ResolveUpgradedLanguage(string semanticKind, string originalLanguage) {
        if (!string.IsNullOrWhiteSpace(originalLanguage)
            && !originalLanguage.Equals("json", StringComparison.OrdinalIgnoreCase)) {
            return originalLanguage;
        }

        return LanguageMode switch {
            MarkdownVisualFenceLanguageMode.GenericSemanticFence => semanticKind,
            MarkdownVisualFenceLanguageMode.IntelligenceXAliasFence => semanticKind switch {
                MarkdownSemanticKinds.Chart => "ix-chart",
                MarkdownSemanticKinds.Network => "ix-network",
                MarkdownSemanticKinds.DataView => "ix-dataview",
                _ => originalLanguage ?? string.Empty
            },
            _ => originalLanguage ?? string.Empty
        };
    }
}

/// <summary>
/// Fence-language strategy for visual code-block upgrades.
/// </summary>
public enum MarkdownVisualFenceLanguageMode {
    /// <summary>Keep the original fence language (typically <c>json</c>).</summary>
    PreserveOriginal = 0,
    /// <summary>Rewrite upgraded payloads to neutral semantic fence languages such as <c>chart</c>.</summary>
    GenericSemanticFence = 1,
    /// <summary>Rewrite upgraded payloads to IntelligenceX alias fences such as <c>ix-chart</c>.</summary>
    IntelligenceXAliasFence = 2
}
