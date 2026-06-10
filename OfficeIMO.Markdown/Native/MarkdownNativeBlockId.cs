namespace OfficeIMO.Markdown;

internal static class MarkdownNativeBlockId {
    internal static string Create(
        MarkdownNativeBlockKind kind,
        IMarkdownBlock sourceBlock,
        MarkdownSyntaxNode syntaxNode,
        MarkdownSourceSpan? sourceSpan) {
        var span = sourceSpan.HasValue ? sourceSpan.Value.ToString() : "nosource";
        var literal = syntaxNode.Literal ?? sourceBlock.RenderMarkdown() ?? string.Empty;
        var key = kind.ToString() + "|" + syntaxNode.Kind + "|" + span + "|" + literal;
        return "mdn-" + ComputeFnv1A64(key).ToString("x16", System.Globalization.CultureInfo.InvariantCulture);
    }

    private static ulong ComputeFnv1A64(string value) {
        const ulong offsetBasis = 14695981039346656037UL;
        const ulong prime = 1099511628211UL;

        var hash = offsetBasis;
        for (var i = 0; i < value.Length; i++) {
            hash ^= value[i];
            hash *= prime;
        }

        return hash;
    }
}
