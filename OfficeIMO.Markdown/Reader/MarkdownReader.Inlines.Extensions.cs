namespace OfficeIMO.Markdown;

internal static partial class MarkdownReader {
    /// <summary>
    /// Parses a single line of Markdown inline content into a typed <see cref="InlineSequence"/>.
    /// This helper is exposed to allow other components (e.g., Word converter) to interpret
    /// inline markup in contexts like table cells where we currently store raw strings.
    /// </summary>
    /// <param name="text">Inline Markdown text.</param>
    /// <param name="options">Reader options controlling inline interpretation.</param>
    /// <returns>Parsed sequence of inline nodes.</returns>
    public static InlineSequence ParseInlineText(string? text, MarkdownLoadOptions? options = null) =>
        ParseInlineText(text, options, null);

    internal static InlineSequence ParseInlineText(string? text, MarkdownLoadOptions? options, MarkdownReaderState? state) =>
        ParseInlineText(text, options, state, sourceMap: null);

    internal static InlineSequence ParseInlineText(
        string? text,
        MarkdownLoadOptions? options,
        MarkdownReaderState? state,
        MarkdownInlineSourceMap? sourceMap) =>
        ParseInlines(text ?? string.Empty, options ?? new MarkdownLoadOptions(), state, sourceMap);

    private static readonly MarkdownInlineParserExtension[] MathOnlyInlineParserExtensions = {
        new MarkdownInlineParserExtension("OfficeIMO.Math", TryParseMathInline)
    };

    private static IReadOnlyList<MarkdownInlineParserExtension> BuildEffectiveInlineParserExtensions(MarkdownLoadOptions options) {
        if (options.InlineParserExtensions.Count == 0) {
            return options.Math ? MathOnlyInlineParserExtensions : Array.Empty<MarkdownInlineParserExtension>();
        }

        var active = new List<MarkdownInlineParserExtension>(options.InlineParserExtensions.Count + 1);
        if (options.Math) {
            active.Add(MathOnlyInlineParserExtensions[0]);
        }

        for (var i = 0; i < options.InlineParserExtensions.Count; i++) {
            var extension = options.InlineParserExtensions[i];
            if (extension != null && extension.AppliesTo(options)) {
                active.Add(extension);
            }
        }

        return active;
    }

    private static IReadOnlyList<MarkdownInlineTransformExtension> BuildEffectiveInlineTransformExtensions(MarkdownLoadOptions options) {
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

    private static void ApplyInlineTransformExtensions(InlineSequence sequence, string sourceText, MarkdownLoadOptions options, MarkdownReaderState? state) {
        var inlineTransformExtensions = BuildEffectiveInlineTransformExtensions(options);
        if (inlineTransformExtensions.Count == 0) {
            return;
        }

        ApplyInlineTransformExtensions(sequence, sourceText, options, state, inlineTransformExtensions, isNestedSequence: false);
    }

    private static void ApplyInlineTransformExtensions(
        InlineSequence sequence,
        string sourceText,
        MarkdownLoadOptions options,
        MarkdownReaderState? state,
        IReadOnlyList<MarkdownInlineTransformExtension> inlineTransformExtensions,
        bool isNestedSequence) {
        for (var i = 0; i < sequence.Nodes.Count; i++) {
            if (sequence.Nodes[i] is IInlineContainerMarkdownInline container && container.NestedInlines != null) {
                ApplyInlineTransformExtensions(
                    container.NestedInlines,
                    sourceText,
                    options,
                    state,
                    inlineTransformExtensions,
                    isNestedSequence: true);
            }
        }

        var context = new MarkdownInlineTransformContext(sourceText, options, state, isNestedSequence);
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
        MarkdownLoadOptions options,
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

    /// <summary>
    /// Parses <c>$...$</c> and <c>$$...$$</c> math. An opening single <c>$</c> must precede a non-space character;
    /// the closing <c>$</c> must follow a non-space character and must not precede a digit.
    /// </summary>
    private static bool TryParseMathInline(MarkdownInlineParserContext context, out MarkdownInlineParseResult result) {
        result = default;
        string text = context.Text;
        int start = context.Position;
        if (start >= text.Length || text[start] != '$') return false;

        if (start + 1 < text.Length && text[start + 1] == '$') {
            int close = text.IndexOf("$$", start + 2, StringComparison.Ordinal);
            if (close < 0) return false;
            string display = text.Substring(start + 2, close - start - 2);
            if (display.Trim().Length == 0) return false;
            result = CreateMathResult(context, display, isDisplay: true, delimiterLength: 2, consumed: close + 2 - start);
            return true;
        }

        if (start + 1 >= text.Length || char.IsWhiteSpace(text[start + 1])) return false;
        for (int j = start + 1; j < text.Length; j++) {
            char value = text[j];
            if (value == '\\') {
                j++;
                continue;
            }

            if (value != '$') continue;
            if (char.IsWhiteSpace(text[j - 1])) continue;
            if (j + 1 < text.Length && char.IsDigit(text[j + 1])) continue;
            string source = text.Substring(start + 1, j - start - 1);
            result = CreateMathResult(context, source, isDisplay: false, delimiterLength: 1, consumed: j + 1 - start);
            return true;
        }

        return false;
    }

    private static MarkdownInlineParseResult CreateMathResult(
        MarkdownInlineParserContext context,
        string source,
        bool isDisplay,
        int delimiterLength,
        int consumed) {
        var math = new MathInline(source, isDisplay) {
            ContentSourceSpan = context.GetSourceSpan(delimiterLength, source.Length)
        };
        return new MarkdownInlineParseResult(math, consumed);
    }
}
