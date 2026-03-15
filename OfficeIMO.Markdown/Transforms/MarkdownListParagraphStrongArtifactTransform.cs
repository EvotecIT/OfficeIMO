using System.Text.RegularExpressions;

namespace OfficeIMO.Markdown;

/// <summary>
/// Repairs malformed strong-marker artifacts inside already-parsed list-item paragraphs.
/// </summary>
/// <remarks>
/// This transform is intended for recoverable list-item inline artifacts that do not require block-boundary repair.
/// It operates after parse, rewrites only list-item paragraph markdown that still contains literal strong-marker damage,
/// and reparses the affected inline content using the current reader options.
/// </remarks>
public sealed class MarkdownListParagraphStrongArtifactTransform : IMarkdownDocumentTransform {
    private static readonly Regex RepeatedStrongDelimiterRunRegex = new(
        @"(?<!\*)(?<left>\*{4,})(?<inner>[^*\r\n]+?)(?<right>\*{4,})(?!\*)",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex DanglingTrailingStrongTokenRegex = new(
        @"(?<token>[\p{L}\p{N}_./:-]+)\*{4}(?<tail>\s*)$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex EscapedDanglingTrailingStrongTokenRegex = new(
        @"(?<token>[\p{L}\p{N}_./:-]+)(?:\\\*){4}(?<tail>\s*)$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex OveropenedMetricValueStrongRegex = new(
        @"^(?<prefix>.*\s)\*{4,}(?<value>[^\s*\r\n][^*\r\n]*?)\*{2}(?<tail>\s*)$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex EscapedOveropenedMetricValueStrongRegex = new(
        @"^(?<prefix>.*\s)(?:\\\*){4,}\*{2}(?<value>[^\s*\r\n][^*\r\n]*?)\*{2}(?<tail>\s*)$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex AdjacentMetricStrongValueRegex = new(
        @"^(?<prefix>.*\s)\*\*(?<first>[^*\r\n]+)\*\*\*{2}(?<second>[^\s*\r\n][^*\r\n]*?)\*{2}(?<tail>\s*)$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex EscapedAdjacentMetricStrongValueRegex = new(
        @"^(?<prefix>.*\s)\\\*\\\*(?<first>[^*\r\n]+)\\\*\\\*\*{2}(?<second>[^\s*\r\n][^*\r\n]*?)\*{2}(?<tail>\s*)$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex MissingTrailingStrongMetricCloseRegex = new(
        @"^(?<prefix>.*\s)\*\*(?<value>[^\r\n*][^\r\n]*?)(?<!\*)\*(?<tail>\s*)$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex EscapedMissingTrailingStrongMetricCloseRegex = new(
        @"^(?<prefix>.*\s)\\\*\*(?<value>[^\r\n*][^\r\n]*?)(?<!\*)\*(?<tail>\s*)$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    /// <summary>
    /// Creates a transform with the specified normalization options.
    /// </summary>
    public MarkdownListParagraphStrongArtifactTransform(MarkdownInputNormalizationOptions options) {
        Options = options ?? throw new ArgumentNullException(nameof(options));
    }

    /// <summary>
    /// Normalization switches used by this transform.
    /// </summary>
    public MarkdownInputNormalizationOptions Options { get; }

    /// <inheritdoc />
    public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        if (context == null) {
            throw new ArgumentNullException(nameof(context));
        }

        MarkdownDocumentBlockRewriter.RewriteDocument(document, block => RewriteBlock(block, context));
        return document;
    }

    private IMarkdownBlock RewriteBlock(IMarkdownBlock block, MarkdownDocumentTransformContext context) {
        switch (block) {
            case OrderedListBlock ordered:
                NormalizeListItems(ordered.Items, context);
                break;
            case UnorderedListBlock unordered:
                NormalizeListItems(unordered.Items, context);
                break;
        }

        return block;
    }

    private void NormalizeListItems(IList<ListItem> items, MarkdownDocumentTransformContext context) {
        for (var i = 0; i < items.Count; i++) {
            var item = items[i];
            if (item == null) {
                continue;
            }

            NormalizeSequence(item.Content, context);
            for (var paragraphIndex = 0; paragraphIndex < item.AdditionalParagraphs.Count; paragraphIndex++) {
                NormalizeSequence(item.AdditionalParagraphs[paragraphIndex], context);
            }

            if (item.Children.Count > 0) {
                NormalizeNestedBlocks(item.Children, context);
            }
        }
    }

    private void NormalizeNestedBlocks(IList<IMarkdownBlock> blocks, MarkdownDocumentTransformContext context) {
        for (var i = 0; i < blocks.Count; i++) {
            var block = blocks[i];
            if (block == null) {
                continue;
            }

            blocks[i] = RewriteBlock(block, context);
        }
    }

    private void NormalizeSequence(InlineSequence? sequence, MarkdownDocumentTransformContext context) {
        var markdown = sequence?.RenderMarkdown() ?? string.Empty;
        if (markdown.Length == 0 || markdown.IndexOf('*') < 0) {
            return;
        }

        var normalized = NormalizeListParagraphMarkdown(markdown);
        if (normalized.Equals(markdown, StringComparison.Ordinal)) {
            return;
        }

        var readerOptions = context.ReaderOptions ?? new MarkdownReaderOptions();
        var reparsed = MarkdownReader.ParseInlineText(normalized, readerOptions);
        sequence!.ReplaceItems(reparsed.Nodes);
    }

    private string NormalizeListParagraphMarkdown(string markdown) {
        var current = markdown;

        if (Options.NormalizeLooseStrongDelimiters) {
            current = RepeatedStrongDelimiterRunRegex.Replace(current, static match => {
                var leftLength = match.Groups["left"].Value.Length;
                var rightLength = match.Groups["right"].Value.Length;
                if (leftLength != rightLength || leftLength % 2 != 0) {
                    return match.Value;
                }

                var inner = match.Groups["inner"].Value.Trim();
                return inner.Length == 0 ? match.Value : "**" + inner + "**";
            });
        }

        if (Options.NormalizeDanglingTrailingStrongListClosers) {
            current = DanglingTrailingStrongTokenRegex.Replace(current, static match => {
                var token = match.Groups["token"].Value.Trim();
                if (token.Length == 0 || token.IndexOf("**", StringComparison.Ordinal) >= 0) {
                    return match.Value;
                }

                return "**" + token + "**" + match.Groups["tail"].Value;
            });

            current = EscapedDanglingTrailingStrongTokenRegex.Replace(current, static match => {
                var token = match.Groups["token"].Value.Trim();
                if (token.Length == 0 || token.IndexOf("**", StringComparison.Ordinal) >= 0) {
                    return match.Value;
                }

                return "**" + token + "**" + match.Groups["tail"].Value;
            });
        }

        if (Options.NormalizeMetricValueStrongRuns) {
            current = OveropenedMetricValueStrongRegex.Replace(current, static match => {
                var value = match.Groups["value"].Value.Trim();
                return value.Length == 0
                    ? match.Value
                    : match.Groups["prefix"].Value + "**" + value + "**" + match.Groups["tail"].Value;
            });

            current = EscapedOveropenedMetricValueStrongRegex.Replace(current, static match => {
                var value = match.Groups["value"].Value.Trim();
                return value.Length == 0
                    ? match.Value
                    : match.Groups["prefix"].Value + "**" + value + "**" + match.Groups["tail"].Value;
            });

            current = AdjacentMetricStrongValueRegex.Replace(current, static match => {
                var first = match.Groups["first"].Value.Trim();
                var second = match.Groups["second"].Value.Trim();
                if (first.Length == 0 || second.Length == 0) {
                    return match.Value;
                }

                if (IsSymbolOnlyValue(first)) {
                    return match.Groups["prefix"].Value + first + " **" + second + "**" + match.Groups["tail"].Value;
                }

                return match.Groups["prefix"].Value
                       + "**"
                       + first
                       + "** **"
                       + second
                       + "**"
                       + match.Groups["tail"].Value;
            });

            current = EscapedAdjacentMetricStrongValueRegex.Replace(current, static match => {
                var first = match.Groups["first"].Value.Trim();
                var second = match.Groups["second"].Value.Trim();
                if (first.Length == 0 || second.Length == 0) {
                    return match.Value;
                }

                if (IsSymbolOnlyValue(first)) {
                    return match.Groups["prefix"].Value + first + " **" + second + "**" + match.Groups["tail"].Value;
                }

                return match.Groups["prefix"].Value
                       + "**"
                       + first
                       + "** **"
                       + second
                       + "**"
                       + match.Groups["tail"].Value;
            });

            current = MissingTrailingStrongMetricCloseRegex.Replace(current, static match => {
                var value = match.Groups["value"].Value.Trim();
                return value.Length == 0
                    ? match.Value
                    : match.Groups["prefix"].Value + "**" + value + "**" + match.Groups["tail"].Value;
            });

            current = EscapedMissingTrailingStrongMetricCloseRegex.Replace(current, static match => {
                var value = match.Groups["value"].Value.Trim();
                return value.Length == 0
                    ? match.Value
                    : match.Groups["prefix"].Value + "**" + value + "**" + match.Groups["tail"].Value;
            });
        }

        return current;
    }

    private static bool IsSymbolOnlyValue(string value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        for (var i = 0; i < value.Length; i++) {
            var ch = value[i];
            if (char.IsWhiteSpace(ch)) {
                continue;
            }

            if (char.IsLetterOrDigit(ch)) {
                return false;
            }
        }

        return true;
    }
}
