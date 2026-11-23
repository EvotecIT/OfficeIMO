using System.Text;

namespace OfficeIMO.Markdown;

/// <summary>
/// One reusable builder per thread for code that builds a string and returns it. A nested caller
/// that finds the builder taken gets a new one, so acquiring is always safe.
/// </summary>
internal static class MarkdownStringBuilderCache {
    private const int MaxRetainedCapacity = 16 * 1024;
    [ThreadStatic] private static StringBuilder? t_cached;

    internal static StringBuilder Acquire(int capacity = 16) {
        StringBuilder? builder = t_cached;
        if (builder != null && capacity <= builder.Capacity) {
            t_cached = null;
            builder.Clear();
            return builder;
        }

        return new StringBuilder(Math.Max(16, capacity));
    }

    internal static string GetStringAndRelease(StringBuilder builder) {
        string value = builder.ToString();
        Release(builder);
        return value;
    }

    internal static void Release(StringBuilder builder) {
        if (builder.Capacity <= MaxRetainedCapacity) {
            t_cached = builder;
        }
    }

    // Whole-document output keeps its own, larger builder so a document write does not grow a
    // fresh builder chunk by chunk.
    private const int MaxRetainedDocumentCapacity = 256 * 1024;
    [ThreadStatic] private static StringBuilder? t_document;

    internal static StringBuilder AcquireDocument() {
        StringBuilder? builder = t_document;
        if (builder == null) {
            return new StringBuilder();
        }

        t_document = null;
        builder.Clear();
        return builder;
    }

    internal static string GetDocumentStringAndRelease(StringBuilder builder) {
        string value = builder.ToString();
        if (builder.Capacity <= MaxRetainedDocumentCapacity) {
            builder.Clear();
            t_document = builder;
        }

        return value;
    }
}
