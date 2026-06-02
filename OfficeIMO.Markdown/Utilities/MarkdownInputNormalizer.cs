using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Text;

namespace OfficeIMO.Markdown;

/// <summary>
/// Stateless markdown input normalizer intended for chat/model outputs before strict parsing.
/// </summary>
public static partial class MarkdownInputNormalizer {
    private const int StrongFlattenMaxIterations = 32;
    private const int LabeledOuterStrongPrefixMaxChars = 120;

    private static readonly Regex ZeroWidthWhitespaceRegex = new Regex(
        @"[\u200B\u2060\uFEFF]",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex EmojiWordJoinRegex = new Regex(
        @"([✅☑✔❌⚠🔥])(?!\s)(?=[\p{L}\p{N}])",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex NumberedChoiceJoinRegex = new Regex(
        @"(\bor|\band|[,;:])(?!\s)(?=\d+\))",
        RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static readonly Regex LetterToNumberedChoiceJoinRegex = new Regex(
        @"(?<=[A-Za-z])(?=\d+\))",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex SentenceCollapsedBulletRegex = new Regex(
        @"(?<=[\.\!\?\)\]])\s*(?=-\s*(?:\*\*[^\r\n]|[A-Z]{2,}\d+\b))",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex InlineCodeSpanRegex = new Regex(
        "`([^`]+)`",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex SoftWrappedStrongRegex = new Regex(
        "\\*\\*(?<left>[^\\r\\n*]{1,80})\\r?\\n(?<right>[^\\r\\n*]{1,80})\\*\\*",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex EscapedInlineCodeSpanRegex = new Regex(
        @"\\`(?<code>[^`\r\n]+?)\\`",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex TightStrongSuffixRegex = new Regex(
        @"(\*\*[^\s*\r\n](?:[^*\r\n]*[^\s*\r\n])?\*\*)(?=[\p{L}\p{N}])",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex TightArrowStrongBoundaryRegex = new Regex(
        @"->\s*(?=\*\*)",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex BrokenStrongArrowLabelRegex = new Regex(
        @"\*\*(?<left>[^*\r\n]{1,200}?)\s*->\s*\*\*(?<label>[^*\r\n]{1,120}?):\*\*",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex WrappedSignalFlowLineRegex = new Regex(
        @"(?m)^(?<prefix>\s*-\s+[^\r\n]*?)\*\*(?<inner>[^\r\n]*->\s*\*\*[^\r\n]*?)\*\*(?<tail>\s*)$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex SignalFlowPlainLabelTightSpacingRegex = new Regex(
        @"^(?<label>[\p{L}][^:\r\n]{0,120}:)(?<next>[^\s/\\])",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex StatusCollapsedLineRegex = new Regex(
        @"(?m)^(?<lead>\s*\*\*Status:[^\r\n]*?)[ \t]-[ \t](?<rest>\*\*.*)$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex BulletCollapsedLineRegex = new Regex(
        @"(?m)^(?<lead>\s*-\s[^\r\n]*?)[ \t]-(?<rest>\s*\*{1,2}.*)$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex LegacyStatusSummaryRegex = new Regex(
        @"(?m)^(?<indent>\s*)\*\*Status:\s*(?<value>[^*\r\n]+)\*\*\s*$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex LegacyBoldMetricBulletRegex = new Regex(
        @"(?m)^(?<indent>\s*-\s)\*\*(?<label>[^*\r\n:]+):\*\*\s*(?<value>[^\r\n]*?)\s*$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex LineStartHostLabelBulletRegex = new Regex(
        @"(?m)^(?<indent>\s*)-(?=[A-Z]{2,}\d+\b)",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex LineStartMissingSpaceBeforeBoldBulletRegex = new Regex(
        @"(?m)^(?<indent>\s*)-(?=\*\*)",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex LineStartUnicodeDashBulletRegex = new Regex(
        @"(?m)^(?<indent>\s*)[‐‑‒–—−](?=(?:\s*\*\*|[A-Z]{2,}\d+\b|[\p{Lu}][\p{L}\p{N}]{1,}\b))",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex LineStartBoldBulletStrongOpenWhitespaceRegex = new Regex(
        @"(?m)^(?<lead>\s*-\s+\*\*)[ \t]+(?=[^\s*\r\n])",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex SingleStarMetricBulletRegex = new Regex(
        @"(?m)^(?<indent>\s*)-\s*\*(?=[A-Za-z][^\r\n]*:\*\*)",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex HeadingListBoundaryRegex = new Regex(
        @"^(?<heading>[ \t]{0,3}#{1,6}[ \t]+[^\r\n]+?)(?<!\s)(?<marker>[-+*])\s+(?=\*\*)",
        RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.Multiline);

    private static readonly Regex CompactStrongLabelListBoundaryRegex = new Regex(
        @"(?<=[\p{P}\p{S}\)])(?<marker>[-+*])\s+(?=\*\*)",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex CompactHeadingBoundaryRegex = new Regex(
        @"(?<=[^\s\r\n])(?<marker>#{2,6})\s+(?=\S)",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex StandaloneSingleHashSeparatorRegex = new Regex(
        @"^\s*#\s*$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex BrokenTwoLineStrongLeadInRegex = new Regex(
        @"^(?<indent>\s*)\*\*(?<label>Result)\s*$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex ColonListBoundaryRegex = new Regex(
        @":\s*(?<marker>[-+*])\s+(?=(\*\*|`|\[|\p{L}|\p{N}))",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly string[] CompactFenceLanguages = {
        "ix-dataview",
        "ix-network",
        "visnetwork",
        "ix-chart",
        "mermaid",
        "network",
        "chart",
        "jsonc",
        "json5",
        "json"
    };

    private static readonly string[] MermaidBodyPrefixes = {
        "flowchart",
        "graph",
        "sequencediagram",
        "classdiagram",
        "statediagram-v2",
        "statediagram",
        "erdiagram",
        "journey",
        "gantt",
        "pie",
        "mindmap",
        "timeline",
        "quadrantchart",
        "xychart",
        "sankey-beta",
        "requirement",
        "gitgraph",
        "c4context",
        "c4container",
        "c4component",
        "c4dynamic",
        "c4deployment"
    };

    private static readonly Regex LooseStrongDelimiterWhitespaceRegex = new Regex(
        @"\*\*(?<inner>[^*\r\n]+)\*\*",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex RepeatedStrongDelimiterRunRegex = new Regex(
        @"(?<!\*)(?<left>\*{4,})(?<inner>[^*\r\n]+?)(?<right>\*{4,})(?!\*)",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex OrderedListMarkerMissingSpaceRegex = new Regex(
        @"^(?<prefix>[ \t]{0,3}\d+[.)])(?=[*_`\[])",
        RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.Multiline);

    private static readonly Regex OrderedListParenMarkerRegex = new Regex(
        @"^(?<indent>[ \t]{0,3})(?<num>\d+)\)\s*(?=\S)",
        RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.Multiline);

    private static readonly Regex OrderedListCaretArtifactRegex = new Regex(
        @"^(?<lead>[ \t]{0,3}\d+\.)\s*\^\s*",
        RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.Multiline);

    private static readonly Regex CollapsedOrderedListAfterParenRegex = new Regex(
        @"(?<=\))\s*(?=\d+\.(?:\^\s*|\s*[*_]{2}|\s+)\S)",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex CollapsedOrderedListAfterDetailRegex = new Regex(
        @"(?<=\))\s+(?=\d+[.)]\s*[*_]{0,2}\s*\S)",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex CollapsedOrderedListAfterStrongRegex = new Regex(
        @"(?<=\*\*)\s+(?=\d+[.)]\s*[*_]{0,2}\s*\S)",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex OrderedItemStrongMissingCloseBeforeParenRegex = new Regex(
        @"(?m)^(?<lead>\s*\d+\.\s+)\*\*(?<title>[^*\r\n()]+)\((?<detail>[^)\r\n]+)\)\s*$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex TightParentheticalSpacingRegex = new Regex(
        @"(?:(?<=\*\*)|(?<=[\p{L}\p{N}\)]))\((?=[\p{L}][^\r\n)]*\))",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex NestedStrongSpanRegex = new Regex(
        @"(?<!\S)\*\*(?<left>[^*\r\n]{6,}?\s)\*\*(?<inner>[A-Za-z0-9`][^*:\r\n]*?)\*\*(?<right>[^*\r\n]*?)\*\*",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex SimpleNestedStrongSpanRegex = new Regex(
        @"\*\*(?<inner>[^*\r\n]+)\*\*",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex LabeledOuterStrongLineRegex = new Regex(
        @"(?m)^(?<prefix>\s*-\s+[^*\r\n]{2," + LabeledOuterStrongPrefixMaxChars.ToString() + @"}\s+\*\*)(?<body>[^\r\n]*)(?<suffix>\*\*)(?<tail>\s*)$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex OrderedListLeadRegex = new Regex(
        @"^\d+[.)]\s+",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex StandaloneHostLabelBulletRegex = new Regex(
        @"^\s*-(?:\s*\*\*)?\s*[A-Z]{2,}\d+(?:\s*\*\*)?\s*:?\s*$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex StructuralMarkdownLineRegex = new Regex(
        @"^(?:[-+*]\s+|\d+[.)]\s+|#{1,6}\s+|>\s?|```|~~~|\|)",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex TrailingDanglingStrongListTokenRegex = new Regex(
        @"(?<token>[\p{L}\p{N}_./:-]+)\*{4}(?<tail>\s*)$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex OveropenedMetricValueStrongRegex = new Regex(
        @"^(?<prefix>\s*(?:-\s+|\d+\.\s+)[^\r\n*]+?\s)\*{4,}(?<value>[^\s*\r\n][^*\r\n]*?)\*{2}(?<tail>\s*)$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex AdjacentMetricStrongValueRegex = new Regex(
        @"^(?<prefix>\s*(?:-\s+|\d+\.\s+)[^\r\n*]+?\s)\*\*(?<first>[^*\r\n]+)\*\*\*{2}(?<second>[^\s*\r\n][^*\r\n]*?)\*{2}(?<tail>\s*)$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex MissingTrailingStrongMetricCloseRegex = new Regex(
        @"^(?<prefix>\s*(?:-\s+|\d+\.\s+)[^\r\n*]+?\s)\*\*(?<value>[^\r\n*][^\r\n]*?)(?<!\*)\*(?<tail>\s*)$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    /// <summary>
    /// Normalizes markdown text based on <paramref name="options"/>.
    /// </summary>
    /// <param name="markdown">Input markdown.</param>
    /// <param name="options">Normalization options.</param>
    /// <returns>Normalized markdown.</returns>
    public static string Normalize(string? markdown, MarkdownInputNormalizationOptions? options = null) {
        var value = markdown ?? string.Empty;
        if (value.Length == 0) {
            return value;
        }

        options ??= new MarkdownInputNormalizationOptions();

        if (options.NormalizeZeroWidthSpacingArtifacts) {
            value = ZeroWidthWhitespaceRegex.Replace(value, string.Empty);
        }

        if (options.NormalizeBrokenTwoLineStrongLeadIns) {
            value = ApplyTransformOutsideFencedCodeBlocks(value, RepairBrokenTwoLineStrongLeadIns);
        }

        if (options.NormalizeSoftWrappedStrongSpans) {
            value = ApplyRegexOutsideFencedCodeBlocks(value, SoftWrappedStrongRegex, static match => {
                var left = match.Groups["left"].Value.Trim();
                var right = match.Groups["right"].Value.Trim();
                if (left.Length == 0 || right.Length == 0) {
                    return match.Value;
                }

                // Avoid collapsing list boundaries such as:
                // **First item text**
                // 2.** Second item**
                if (LooksLikeListMarkerFragment(right)) {
                    return match.Value;
                }

                return "**" + left + " " + right + "**";
            });
        }

        if (options.NormalizeWrappedSignalFlowStrongRuns) {
            value = ApplyRegexOutsideFencedCodeBlocks(
                value,
                WrappedSignalFlowLineRegex,
                static match => {
                    var inner = match.Groups["inner"].Value;
                    var markerIndex = inner.IndexOf("-> **", StringComparison.Ordinal);
                    if (markerIndex < 0) {
                        markerIndex = inner.IndexOf("->**", StringComparison.Ordinal);
                    }

                    if (markerIndex <= 0) {
                        return match.Value;
                    }

                    var headline = inner.Substring(0, markerIndex).TrimEnd();
                    if (headline.Length == 0) {
                        return match.Value;
                    }

                    var flow = inner.Substring(markerIndex).TrimStart();
                    if (flow.StartsWith("->**", StringComparison.Ordinal)) {
                        flow = "-> **" + flow.Substring(4);
                    }

                    if (!flow.StartsWith("-> **", StringComparison.Ordinal)) {
                        return match.Value;
                    }

                    return match.Groups["prefix"].Value + "**" + headline + "** " + flow + match.Groups["tail"].Value;
                });
        }

        if (options.NormalizeSignalFlowLabelSpacing) {
            value = ApplyTransformOutsideFencedCodeBlocks(value, NormalizeSignalFlowLabelSpacing);
        }

        if (options.NormalizeEmojiWordJoins) {
            value = ApplyRegexOutsideFencedCodeBlocks(value, EmojiWordJoinRegex, static match => match.Groups[1].Value + " ");
        }

        if (options.NormalizeCompactNumberedChoiceBoundaries) {
            value = ApplyRegexOutsideFencedCodeBlocks(value, NumberedChoiceJoinRegex, static match => match.Groups[1].Value + " ");
            value = ApplyRegexOutsideFencedCodeBlocks(value, LetterToNumberedChoiceJoinRegex, static _ => " ");
        }

        if (options.NormalizeSentenceCollapsedBullets) {
            value = ApplyRegexOutsideFencedCodeBlocks(value, SentenceCollapsedBulletRegex, static _ => "\n", preserveInlineCodeSpans: true);
        }

        if (options.NormalizeCollapsedMetricChains) {
            value = ExpandCollapsedMetricLines(value);
            value = NormalizeLegacyMetricBulletLeads(value);
            value = ConvertLegacyMetricMarkdown(value);
        }

        if (options.NormalizeHostLabelBulletArtifacts) {
            value = ApplyTransformOutsideFencedCodeBlocks(value, NormalizeHostLabelBulletArtifacts);
        }

        if (options.NormalizeNestedStrongDelimiters) {
            value = FlattenNestedStrongSpansOutsideFencedCodeBlocks(value);
        }

        if (options.NormalizeLooseStrongDelimiters) {
            value = ApplyRegexOutsideFencedCodeBlocks(value, RepeatedStrongDelimiterRunRegex, static match => {
                var leftLength = match.Groups["left"].Value.Length;
                var rightLength = match.Groups["right"].Value.Length;
                if (leftLength != rightLength || leftLength % 2 != 0) {
                    return match.Value;
                }

                var inner = match.Groups["inner"].Value;
                var trimmed = inner.Trim();
                if (trimmed.Length == 0) {
                    return match.Value;
                }

                return "**" + trimmed + "**";
            });

            value = ApplyRegexOutsideFencedCodeBlocks(value, LooseStrongDelimiterWhitespaceRegex, static match => {
                var inner = match.Groups["inner"].Value;
                var trimmed = inner.Trim();
                if (trimmed.Length == 0 || trimmed.Length == inner.Length) {
                    return match.Value;
                }

                return "**" + trimmed + "**";
            });
        }

        if (options.NormalizeDanglingTrailingStrongListClosers) {
            value = RepairDanglingTrailingStrongListClosers(value);
        }

        if (options.NormalizeMetricValueStrongRuns) {
            value = RepairMalformedMetricValueStrongRuns(value);
        }

        if (options.NormalizeCollapsedOrderedListBoundaries) {
            value = ApplyRegexOutsideFencedCodeBlocks(value, CollapsedOrderedListAfterDetailRegex, static _ => "\n");
            value = ApplyRegexOutsideFencedCodeBlocks(value, CollapsedOrderedListAfterStrongRegex, static _ => "\n");
            value = ApplyRegexOutsideFencedCodeBlocks(value, CollapsedOrderedListAfterParenRegex, static _ => "\n");
        }

        if (options.NormalizeTightStrongBoundaries) {
            value = ApplyRegexOutsideFencedCodeBlocks(value, TightStrongSuffixRegex, static match => match.Groups[1].Value + " ");
        }

        if (options.NormalizeTightArrowStrongBoundaries) {
            value = ApplyRegexOutsideFencedCodeBlocks(
                value,
                TightArrowStrongBoundaryRegex,
                static _ => "-> ",
                preserveInlineCodeSpans: true);
        }

        if (options.NormalizeBrokenStrongArrowLabels) {
            value = ApplyRegexOutsideFencedCodeBlocks(
                value,
                BrokenStrongArrowLabelRegex,
                static match => {
                    var left = match.Groups["left"].Value.Trim();
                    var label = match.Groups["label"].Value.Trim();
                    if (left.Length == 0 || label.Length == 0) {
                        return match.Value;
                    }

                    return "**" + left + "** -> **" + label + ":**";
                },
                preserveInlineCodeSpans: true);
        }

        if (options.NormalizeCompactHeadingBoundaries) {
            value = SplitCollapsedTableHeadingBoundaries(value);
            value = ApplyRegexOutsideFencedCodeBlocks(
                value,
                CompactHeadingBoundaryRegex,
                static match => "\n" + match.Groups["marker"].Value + " ",
                preserveInlineCodeSpans: true);
        }

        if (options.NormalizeStandaloneHashHeadingSeparators) {
            value = RemoveStandaloneHashSeparatorsBeforeHeadings(value);
        }

        if (options.NormalizeHeadingListBoundaries) {
            value = ApplyRegexOutsideFencedCodeBlocks(
                value,
                HeadingListBoundaryRegex,
                static match => match.Groups["heading"].Value.TrimEnd() + "\n" + match.Groups["marker"].Value + " ",
                preserveInlineCodeSpans: true);
        }

        if (options.NormalizeCompactStrongLabelListBoundaries) {
            value = ApplyRegexOutsideFencedCodeBlocks(
                value,
                CompactStrongLabelListBoundaryRegex,
                static match => "\n" + match.Groups["marker"].Value + " ",
                preserveInlineCodeSpans: true);
        }

        if (options.NormalizeColonListBoundaries) {
            value = ApplyRegexOutsideFencedCodeBlocks(
                value,
                ColonListBoundaryRegex,
                static match => ":\n" + match.Groups["marker"].Value + " ",
                preserveInlineCodeSpans: true);
        }

        if (options.NormalizeCompactFenceBodyBoundaries) {
            value = NormalizeCompactFenceBodyBoundaries(value);
        }

        if (options.NormalizeOrderedListMarkerSpacing) {
            value = ApplyRegexOutsideFencedCodeBlocks(value, OrderedListMarkerMissingSpaceRegex, static match => match.Groups["prefix"].Value + " ");
        }

        if (options.NormalizeOrderedListParenMarkers) {
            value = ApplyRegexOutsideFencedCodeBlocks(value, OrderedListParenMarkerRegex, static match => match.Groups["indent"].Value + match.Groups["num"].Value + ". ");
        }

        if (options.NormalizeOrderedListCaretArtifacts) {
            value = ApplyRegexOutsideFencedCodeBlocks(value, OrderedListCaretArtifactRegex, static match => match.Groups["lead"].Value + " ");
        }

        if (options.NormalizeTightParentheticalSpacing) {
            value = ApplyRegexOutsideFencedCodeBlocks(
                value,
                TightParentheticalSpacingRegex,
                static _ => " (",
                preserveInlineCodeSpans: true);
        }

        if (options.NormalizeOrderedListStrongDetailClosures) {
            value = ApplyRegexOutsideFencedCodeBlocks(value, OrderedItemStrongMissingCloseBeforeParenRegex, static match => {
                var lead = match.Groups["lead"].Value;
                var title = match.Groups["title"].Value.Trim();
                var detail = match.Groups["detail"].Value.Trim();
                return lead + "**" + title + "** (" + detail + ")";
            });
        }

        // Keep a final loose-strong pass after ordered-list detail repair because that step can
        // introduce new boundary whitespace inside reconstructed strong delimiters.
        if (options.NormalizeLooseStrongDelimiters) {
            value = ApplyRegexOutsideFencedCodeBlocks(value, RepeatedStrongDelimiterRunRegex, static match => {
                var leftLength = match.Groups["left"].Value.Length;
                var rightLength = match.Groups["right"].Value.Length;
                if (leftLength != rightLength || leftLength % 2 != 0) {
                    return match.Value;
                }

                var inner = match.Groups["inner"].Value;
                var trimmed = inner.Trim();
                if (trimmed.Length == 0) {
                    return match.Value;
                }

                return "**" + trimmed + "**";
            });

            value = ApplyRegexOutsideFencedCodeBlocks(value, LooseStrongDelimiterWhitespaceRegex, static match => {
                var inner = match.Groups["inner"].Value;
                var trimmed = inner.Trim();
                if (trimmed.Length == 0 || trimmed.Length == inner.Length) {
                    return match.Value;
                }

                return "**" + trimmed + "**";
            });
        }

        if (options.NormalizeInlineCodeSpanLineBreaks) {
            value = ApplyRegexOutsideFencedCodeBlocks(value, InlineCodeSpanRegex, static match => {
                var body = match.Groups[1].Value;
                if (body.IndexOfAny(new[] { '\r', '\n' }) < 0) {
                    return match.Value;
                }

                var compact = body.Replace("\r\n", " ")
                    .Replace('\r', ' ')
                    .Replace('\n', ' ')
                    .Trim();
                return compact.Length == 0 ? "``" : "`" + compact + "`";
            });
        }

        if (options.NormalizeEscapedInlineCodeSpans) {
            value = ApplyRegexOutsideFencedCodeBlocks(value, EscapedInlineCodeSpanRegex, static match => {
                var body = match.Groups["code"].Value;
                return body.Length == 0 ? "``" : "`" + body + "`";
            });
        }

        return value;
    }
}
