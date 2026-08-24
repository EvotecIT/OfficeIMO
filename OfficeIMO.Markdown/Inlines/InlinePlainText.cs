namespace OfficeIMO.Markdown;

internal static class InlinePlainText {
    public static string Extract(InlineSequence? sequence) {
        if (sequence == null || sequence.Nodes.Count == 0) {
            return string.Empty;
        }

        if (sequence.Nodes.Count == 1) {
            IMarkdownInline node = sequence.Nodes[0];
            if (node is MarkdownTextRun textRun) {
                return textRun.Text;
            }
            if (node is CodeSpanInline code) {
                return code.Text;
            }
            if (node is IInlineContainerMarkdownInline container && container.NestedInlines != null) {
                return Extract(container.NestedInlines);
            }
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
