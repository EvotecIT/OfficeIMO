using System.Runtime.CompilerServices;

namespace OfficeIMO.Markdown;

internal static class MarkdownInlineSourceSpans {
    private sealed class Holder {
        public MarkdownSourceSpan? Span;
    }

    private static readonly ConditionalWeakTable<IMarkdownInline, Holder> _spans = new();

    internal static MarkdownSourceSpan? Get(IMarkdownInline? inline) {
        if (inline == null) {
            return null;
        }

        if (inline is MarkdownInline builtInInline) {
            return builtInInline.SourceSpan;
        }

        return _spans.TryGetValue(inline, out var holder) ? holder.Span : null;
    }

    internal static void Set(IMarkdownInline? inline, MarkdownSourceSpan? span) {
        if (inline == null || span == null) {
            return;
        }

        if (inline is MarkdownInline builtInInline) {
            builtInInline.SourceSpan = span;
            return;
        }

        _spans.Remove(inline);
        _spans.Add(inline, new Holder {
            Span = span
        });
    }
}
