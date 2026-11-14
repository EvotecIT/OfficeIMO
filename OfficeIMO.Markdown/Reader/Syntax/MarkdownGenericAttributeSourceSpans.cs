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

    // A parse without a syntax tree has no span to record, but a code block still has to know
    // that its attributes came from a standalone attribute line.
    private static readonly ConditionalWeakTable<MarkdownObject, object> _standaloneWithoutSpan = new();
    private static readonly object StandaloneMarker = new();

    internal static void MarkStandaloneWithoutSpan(MarkdownObject? markdownObject) {
        if (markdownObject != null) {
            _standaloneWithoutSpan.GetValue(markdownObject, static _ => StandaloneMarker);
        }
    }

    internal static bool HasStandaloneSource(MarkdownObject? markdownObject) =>
        GetSourceSpan(markdownObject).HasValue
        || (markdownObject != null && _standaloneWithoutSpan.TryGetValue(markdownObject, out _));
}
