namespace OfficeIMO.Markdown;

internal static class MarkdownNativeInlineId {
    internal static string Create(
        MarkdownNativeInlineKind kind,
        MarkdownSyntaxNode syntaxNode,
        MarkdownSourceSpan? sourceSpan) {
        var span = sourceSpan.HasValue ? sourceSpan.Value.ToString() : "nosource";
        var literal = syntaxNode.Literal ?? string.Empty;
        var key = kind.ToString() + "|" + syntaxNode.Kind + "|" + span + "|" + literal;
        return "mdn-in-" + ComputeFnv1A64(key).ToString("x16", System.Globalization.CultureInfo.InvariantCulture);
    }

    internal static ulong ComputeFnv1A64(string value) {
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
