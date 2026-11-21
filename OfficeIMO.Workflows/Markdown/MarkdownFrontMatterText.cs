using OfficeIMO.Markdown;

namespace OfficeIMO.Workflows;

internal static class MarkdownFrontMatterText {
    /// <summary>Returns a top-level front matter scalar's text, or <see langword="null"/> when it is missing, null or a collection.</summary>
    internal static string? Get(MarkdownDocument document, string key) =>
        document.TryGetFrontMatterValue(key, out string? value) ? value : null;
}
