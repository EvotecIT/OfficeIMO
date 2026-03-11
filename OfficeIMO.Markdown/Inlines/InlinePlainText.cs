namespace OfficeIMO.Markdown;

internal static class InlinePlainText {
    public static string Extract(InlineSequence? sequence) {
        if (sequence == null || sequence.Nodes.Count == 0) {
            return string.Empty;
        }

        var sb = new System.Text.StringBuilder();
        AppendPlainText(sb, sequence);
        return sb.ToString();
    }

    internal static void AppendPlainText(System.Text.StringBuilder sb, InlineSequence sequence) {
        foreach (var node in sequence.Nodes) {
            GetPlainTextNode(node).AppendPlainText(sb);
        }
    }

    private static IPlainTextMarkdownInline GetPlainTextNode(IMarkdownInline node) {
        return node as IPlainTextMarkdownInline
            ?? throw new InvalidOperationException($"Inline node of type '{node.GetType().FullName}' does not implement {nameof(IPlainTextMarkdownInline)}.");
    }
}
