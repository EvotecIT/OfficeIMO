using OfficeIMO.Markdown;

namespace OfficeIMO.Word.Markdown;

/// <summary>
/// Shared fenced-block contract for Word-specific semantic markdown exported by <c>OfficeIMO.Word.Markdown</c>.
/// </summary>
public static class WordMarkdownSemanticBlocks {
    /// <summary>Semantic kind used for Word header payloads.</summary>
    public const string HeaderSemanticKind = "word-header";

    /// <summary>Semantic kind used for Word footer payloads.</summary>
    public const string FooterSemanticKind = "word-footer";

    /// <summary>Fence language used for Word header payloads.</summary>
    public const string HeaderFenceLanguage = "officeimo-word-header";

    /// <summary>Fence language used for Word footer payloads.</summary>
    public const string FooterFenceLanguage = "officeimo-word-footer";

    /// <summary>
    /// Creates reader options with the Word header/footer semantic fenced-block extensions registered.
    /// </summary>
    public static MarkdownReaderOptions CreateReaderOptions(
        MarkdownReaderOptions.MarkdownDialectProfile profile = MarkdownReaderOptions.MarkdownDialectProfile.OfficeIMO) {
        var options = MarkdownReaderOptions.CreateProfile(profile);
        ConfigureReaderOptions(options);
        return options;
    }

    /// <summary>
    /// Registers Word header/footer fenced-block extensions on an existing reader options instance.
    /// </summary>
    public static void ConfigureReaderOptions(MarkdownReaderOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        AddExtensionIfMissing(
            options,
            HeaderFenceLanguage,
            HeaderSemanticKind,
            "Word header semantic block");

        AddExtensionIfMissing(
            options,
            FooterFenceLanguage,
            FooterSemanticKind,
            "Word footer semantic block");
    }

    private static void AddExtensionIfMissing(
        MarkdownReaderOptions options,
        string language,
        string semanticKind,
        string name) {
        for (int i = 0; i < options.FencedBlockExtensions.Count; i++) {
            var extension = options.FencedBlockExtensions[i];
            for (int j = 0; j < extension.Languages.Count; j++) {
                if (string.Equals(extension.Languages[j], language, StringComparison.OrdinalIgnoreCase)) {
                    return;
                }
            }
        }

        options.FencedBlockExtensions.Add(new MarkdownFencedBlockExtension(
            name,
            new[] { language },
            context => new SemanticFencedBlock(semanticKind, context.InfoString, context.Content, context.Caption)));
    }
}
