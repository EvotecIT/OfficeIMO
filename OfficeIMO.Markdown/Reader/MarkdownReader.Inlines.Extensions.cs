namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    /// <summary>
    /// Parses a single line of Markdown inline content into a typed <see cref="InlineSequence"/>.
    /// This helper is exposed to allow other components (e.g., Word converter) to interpret
    /// inline markup in contexts like table cells where we currently store raw strings.
    /// </summary>
    /// <param name="text">Inline Markdown text.</param>
    /// <param name="options">Reader options controlling inline interpretation.</param>
    /// <returns>Parsed sequence of inline nodes.</returns>
    public static InlineSequence ParseInlineText(string? text, MarkdownReaderOptions? options = null) =>
        ParseInlineText(text, options, null);

    internal static InlineSequence ParseInlineText(string? text, MarkdownReaderOptions? options, MarkdownReaderState? state) =>
        ParseInlineText(text, options, state, sourceMap: null);

    internal static InlineSequence ParseInlineText(
        string? text,
        MarkdownReaderOptions? options,
        MarkdownReaderState? state,
        MarkdownInlineSourceMap? sourceMap) =>
        ParseInlines(text ?? string.Empty, options ?? new MarkdownReaderOptions(), state, sourceMap);

    private static IReadOnlyList<MarkdownInlineParserExtension> BuildEffectiveInlineParserExtensions(MarkdownReaderOptions options) {
        if (options.InlineParserExtensions.Count == 0) {
            return Array.Empty<MarkdownInlineParserExtension>();
        }

        var active = new List<MarkdownInlineParserExtension>(options.InlineParserExtensions.Count);
        for (var i = 0; i < options.InlineParserExtensions.Count; i++) {
            var extension = options.InlineParserExtensions[i];
            if (extension != null && extension.AppliesTo(options)) {
                active.Add(extension);
            }
        }

        return active;
    }

    private static IReadOnlyList<MarkdownInlineTransformExtension> BuildEffectiveInlineTransformExtensions(MarkdownReaderOptions options) {
        if (options.InlineTransformExtensions.Count == 0) {
            return Array.Empty<MarkdownInlineTransformExtension>();
        }

        var active = new List<MarkdownInlineTransformExtension>(options.InlineTransformExtensions.Count);
        for (var i = 0; i < options.InlineTransformExtensions.Count; i++) {
            var extension = options.InlineTransformExtensions[i];
            if (extension != null && extension.AppliesTo(options)) {
                active.Add(extension);
            }
        }

        return active;
    }

    private static void ApplyInlineTransformExtensions(InlineSequence sequence, string sourceText, MarkdownReaderOptions options) {
        var inlineTransformExtensions = BuildEffectiveInlineTransformExtensions(options);
        if (inlineTransformExtensions.Count == 0) {
            return;
        }

        ApplyInlineTransformExtensions(sequence, sourceText, options, inlineTransformExtensions, isNestedSequence: false);
    }

    private static void ApplyInlineTransformExtensions(
        InlineSequence sequence,
        string sourceText,
        MarkdownReaderOptions options,
        IReadOnlyList<MarkdownInlineTransformExtension> inlineTransformExtensions,
        bool isNestedSequence) {
        for (var i = 0; i < sequence.Nodes.Count; i++) {
            if (sequence.Nodes[i] is IInlineContainerMarkdownInline container && container.NestedInlines != null) {
                ApplyInlineTransformExtensions(
                    container.NestedInlines,
                    sourceText,
                    options,
                    inlineTransformExtensions,
                    isNestedSequence: true);
            }
        }

        var context = new MarkdownInlineTransformContext(sourceText, options, isNestedSequence);
        for (var i = 0; i < inlineTransformExtensions.Count; i++) {
            var extension = inlineTransformExtensions[i];
            var transformed = extension.Transform(sequence, context);
            if (transformed == null || ReferenceEquals(transformed, sequence)) {
                continue;
            }

            sequence.ReplaceItems(transformed.Nodes);
        }
    }

    private static bool TryParseInlineExtension(
        string text,
        int position,
        MarkdownReaderOptions options,
        MarkdownReaderState? state,
        bool allowLinks,
        bool allowImages,
        MarkdownInlineSourceMap? sourceMap,
        IReadOnlyList<MarkdownInlineParserExtension> inlineParserExtensions,
        Func<int, int, bool, bool, InlineSequence> parseNestedInlineSegment,
        out MarkdownInlineParseResult result) {
        result = default;
        if (inlineParserExtensions.Count == 0) {
            return false;
        }

        var context = new MarkdownInlineParserContext(
            text,
            position,
            options,
            state,
            allowLinks,
            allowImages,
            sourceMap,
            parseNestedInlineSegment);

        for (var i = 0; i < inlineParserExtensions.Count; i++) {
            var extension = inlineParserExtensions[i];
            if (!extension.Parser(context, out result)) {
                continue;
            }

            if (result.ConsumedLength <= 0) {
                throw new InvalidOperationException($"Inline parser extension '{extension.Name}' returned a non-positive consumed length.");
            }

            if (position + result.ConsumedLength > text.Length) {
                throw new InvalidOperationException($"Inline parser extension '{extension.Name}' consumed past the end of the input.");
            }

            return true;
        }

        result = default;
        return false;
    }
}
