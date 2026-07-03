using System.Runtime.CompilerServices;

namespace OfficeIMO.Markdown;

internal static class MarkdownGenericAttributeSourceSpans {
    private sealed class Holder {
        public string SourceText = string.Empty;
        public MarkdownSourceSpan? SourceSpan;
    }

    private static readonly ConditionalWeakTable<MarkdownObject, Holder> _spans = new();

    internal static void Set(MarkdownObject? markdownObject, string? sourceText, MarkdownSourceSpan? sourceSpan) {
        if (markdownObject == null || string.IsNullOrEmpty(sourceText) || !sourceSpan.HasValue) {
            return;
        }

        var holder = _spans.GetValue(markdownObject, static _ => new Holder());
        holder.SourceText = sourceText ?? string.Empty;
        holder.SourceSpan = sourceSpan;
    }

    internal static string? GetSourceText(MarkdownObject? markdownObject) =>
        markdownObject != null && _spans.TryGetValue(markdownObject, out var holder)
            ? holder.SourceText
            : null;

    internal static MarkdownSourceSpan? GetSourceSpan(MarkdownObject? markdownObject) =>
        markdownObject != null && _spans.TryGetValue(markdownObject, out var holder)
            ? holder.SourceSpan
            : null;
}
