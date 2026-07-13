namespace OfficeIMO.Markdown.Pdf;

/// <summary>Registers Markdown fenced-block semantics understood by the PDF adapter.</summary>
public static class MarkdownPdfSemanticBlocks {
    /// <summary>Creates Markdown reader options with the PDF visual fenced-block mappings registered.</summary>
    public static MarkdownReaderOptions CreateReaderOptions(
        MarkdownReaderOptions.MarkdownDialectProfile profile = MarkdownReaderOptions.MarkdownDialectProfile.OfficeIMO) {
        MarkdownReaderOptions options = MarkdownReaderOptions.CreateProfile(profile);
        ConfigureReaderOptions(options);
        return options;
    }

    /// <summary>Registers the PDF visual fenced-block mappings on existing reader options.</summary>
    public static void ConfigureReaderOptions(MarkdownReaderOptions options) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        options.FencedBlockExtensions.Add(new MarkdownFencedBlockExtension(
            "OfficeIMO Markdown PDF visual fences",
            new[] { "chart", "ix-chart", "mermaid", "network", "visnetwork", "ix-network", "dataview", "ix-dataview" },
            static context => new SemanticFencedBlock(
                ResolveSemanticKind(context.Language),
                context.InfoString,
                context.Content,
                context.Caption)));
    }

    private static string ResolveSemanticKind(string? language) {
        switch ((language ?? string.Empty).Trim().ToLowerInvariant()) {
            case "chart":
            case "ix-chart":
                return MarkdownSemanticKinds.Chart;
            case "mermaid":
                return MarkdownSemanticKinds.Mermaid;
            case "network":
            case "visnetwork":
            case "ix-network":
                return MarkdownSemanticKinds.Network;
            case "dataview":
            case "ix-dataview":
                return MarkdownSemanticKinds.DataView;
            default:
                return MarkdownSemanticKinds.Custom;
        }
    }
}
