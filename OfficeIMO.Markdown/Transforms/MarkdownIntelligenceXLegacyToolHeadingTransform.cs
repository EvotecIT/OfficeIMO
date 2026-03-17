using System.Text.RegularExpressions;

namespace OfficeIMO.Markdown;

/// <summary>
/// Promotes parseable legacy IntelligenceX tool-heading artifacts into real heading blocks.
/// </summary>
/// <remarks>
/// This transform intentionally targets only cases where the markdown already parsed into a recoverable AST,
/// so transcript compatibility can move away from text rewriting and toward structural upgrades.
/// </remarks>
public sealed class MarkdownIntelligenceXLegacyToolHeadingTransform : IMarkdownDocumentTransform {
    private static readonly Regex LegacyToolHeadingBulletRegex = new(
        @"^(?<tool>[a-z0-9_.-]+):\s*(?<heading>#{2,6}\s+[^\r\n]+)\s*$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    private static readonly Regex LegacyToolHeadingLeadRegex = new(
        @"^(?<tool>[a-z0-9_.-]+):\s*$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    private static readonly Regex LegacyToolHeadingSplitBulletRegex = new(
        @"^(?<tool>[a-z0-9_.-]+):\s*(?<fragment>#{1,5})\s*$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    private static readonly Regex LegacyToolSlugHeadingRegex = new(
        @"^[a-z0-9_.-]+$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    private static readonly Regex MarkdownHeadingRegex = new(
        @"^(?<level>#{2,6})\s+(?<text>.+?)\s*$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    /// <inheritdoc />
    public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        if (context == null) {
            throw new ArgumentNullException(nameof(context));
        }

        MarkdownDocumentBlockListExpander.RewriteDocument(document, context, RewriteBlocks);
        return document;
    }

    private static List<IMarkdownBlock> RewriteBlocks(
        IReadOnlyList<IMarkdownBlock> blocks,
        MarkdownDocumentTransformContext context) {
        var rewritten = new List<IMarkdownBlock>(blocks.Count);
        for (var i = 0; i < blocks.Count; i++) {
            var block = blocks[i];
            if (block == null) {
                continue;
            }

            if (block is HeadingBlock slugHeading
                && i + 1 < blocks.Count
                && blocks[i + 1] is HeadingBlock
                && IsLegacyToolSlugHeading(slugHeading)) {
                continue;
            }

            if (block is UnorderedListBlock unordered
                && TryRewriteLegacyToolHeadingList(
                    unordered,
                    i + 1 < blocks.Count ? blocks[i + 1] as HeadingBlock : null,
                    context,
                    out var expandedBlocks,
                    out var consumedFollowingHeading)) {
                rewritten.AddRange(expandedBlocks);
                if (consumedFollowingHeading) {
                    i++;
                }
                continue;
            }

            rewritten.Add(block);
        }

        return rewritten;
    }

    private static bool TryRewriteLegacyToolHeadingList(
        UnorderedListBlock list,
        HeadingBlock? followingHeading,
        MarkdownDocumentTransformContext context,
        out IReadOnlyList<IMarkdownBlock> blocks,
        out bool consumedFollowingHeading) {
        var rewritten = new List<IMarkdownBlock>();
        var retainedItems = new List<ListItem>();
        var changed = false;
        consumedFollowingHeading = false;

        for (var i = 0; i < list.Items.Count; i++) {
            var item = list.Items[i];
            if (item == null) {
                continue;
            }

            var nextHeadingForItem = i == list.Items.Count - 1 ? followingHeading : null;
            if (TryPromoteLegacyToolHeadingItem(item, nextHeadingForItem, context, out var promotedBlocks, out var consumedNextHeadingForItem)) {
                FlushRetainedItems(retainedItems, rewritten);
                rewritten.AddRange(promotedBlocks);
                changed = true;
                consumedFollowingHeading |= consumedNextHeadingForItem;
                continue;
            }

            retainedItems.Add(item);
        }

        if (!changed) {
            blocks = Array.Empty<IMarkdownBlock>();
            return false;
        }

        FlushRetainedItems(retainedItems, rewritten);
        blocks = rewritten;
        return true;
    }

    private static bool TryPromoteLegacyToolHeadingItem(
        ListItem item,
        HeadingBlock? followingHeading,
        MarkdownDocumentTransformContext context,
        out IReadOnlyList<IMarkdownBlock> blocks,
        out bool consumedFollowingHeading) {
        blocks = Array.Empty<IMarkdownBlock>();
        consumedFollowingHeading = false;
        if (item.IsTask) {
            return false;
        }

        var itemBlocks = item.BlockChildren;
        if (itemBlocks.Count == 0 || itemBlocks[0] is not ParagraphBlock firstParagraph) {
            return false;
        }

        var paragraphMarkdown = InlinePlainText.Extract(firstParagraph.Inlines).Trim();
        if (paragraphMarkdown.Length == 0) {
            return false;
        }

        var trailingBlocks = new List<IMarkdownBlock>(Math.Max(0, itemBlocks.Count - 1));
        for (var i = 1; i < itemBlocks.Count; i++) {
            if (itemBlocks[i] != null) {
                trailingBlocks.Add(itemBlocks[i]);
            }
        }

        if (LegacyToolHeadingBulletRegex.Match(paragraphMarkdown) is { Success: true } inlineMatch
            && TryCreateHeadingBlock(inlineMatch.Groups["heading"].Value, context.ReaderOptions, out var inlineHeading)) {
            blocks = BuildPromotedBlocks(inlineHeading, trailingBlocks, skipFirstPromotedBlock: false);
            return true;
        }

        if (trailingBlocks.Count > 0 && TryResolveTrailingHeadingCandidate(trailingBlocks[0], context.ReaderOptions, out var promotedHeading, out var consumedTrailingHeadingBlock)) {
            if (LegacyToolHeadingLeadRegex.IsMatch(paragraphMarkdown)) {
                blocks = BuildPromotedBlocks(promotedHeading, trailingBlocks, skipFirstPromotedBlock: consumedTrailingHeadingBlock);
                return true;
            }

            if (LegacyToolHeadingSplitBulletRegex.Match(paragraphMarkdown) is not { Success: true } splitChildMatch) {
                return false;
            }

            var combinedChildLevel = Math.Min(6, splitChildMatch.Groups["fragment"].Value.Length + promotedHeading.Level);
            var combinedChildHeading = new HeadingBlock(combinedChildLevel, promotedHeading.Inlines);
            blocks = BuildPromotedBlocks(combinedChildHeading, trailingBlocks, skipFirstPromotedBlock: consumedTrailingHeadingBlock);
            return true;
        }

        if (followingHeading == null) {
            return false;
        }

        if (LegacyToolHeadingLeadRegex.IsMatch(paragraphMarkdown)) {
            blocks = new IMarkdownBlock[] { followingHeading };
            consumedFollowingHeading = true;
            return true;
        }

        var splitMatch = LegacyToolHeadingSplitBulletRegex.Match(paragraphMarkdown);
        if (!splitMatch.Success) {
            return false;
        }

        var combinedLevel = Math.Min(6, splitMatch.Groups["fragment"].Value.Length + followingHeading.Level);
        blocks = new IMarkdownBlock[] { new HeadingBlock(combinedLevel, followingHeading.Inlines) };
        consumedFollowingHeading = true;
        return true;
    }

    private static IReadOnlyList<IMarkdownBlock> BuildPromotedBlocks(
        HeadingBlock heading,
        IReadOnlyList<IMarkdownBlock> childBlocks,
        bool skipFirstPromotedBlock) {
        var blocks = new List<IMarkdownBlock> {
            heading
        };

        for (var i = skipFirstPromotedBlock ? 1 : 0; i < childBlocks.Count; i++) {
            if (childBlocks[i] != null) {
                blocks.Add(childBlocks[i]);
            }
        }

        return blocks;
    }

    private static bool TryResolveTrailingHeadingCandidate(
        IMarkdownBlock block,
        MarkdownReaderOptions? readerOptions,
        out HeadingBlock heading,
        out bool consumedOriginalBlock) {
        if (block is HeadingBlock headingBlock) {
            heading = headingBlock;
            consumedOriginalBlock = true;
            return true;
        }

        if (block is ParagraphBlock paragraph
            && TryCreateHeadingBlock(((IMarkdownBlock)paragraph).RenderMarkdown(), readerOptions, out var paragraphHeading)) {
            heading = paragraphHeading;
            consumedOriginalBlock = true;
            return true;
        }

        heading = null!;
        consumedOriginalBlock = false;
        return false;
    }

    private static void FlushRetainedItems(List<ListItem> retainedItems, List<IMarkdownBlock> target) {
        if (retainedItems.Count == 0) {
            return;
        }

        var list = new UnorderedListBlock();
        for (var i = 0; i < retainedItems.Count; i++) {
            list.Items.Add(retainedItems[i]);
        }

        target.Add(list);
        retainedItems.Clear();
    }

    private static bool TryCreateHeadingBlock(
        string headingMarkdown,
        MarkdownReaderOptions? readerOptions,
        out HeadingBlock heading) {
        heading = null!;
        var match = MarkdownHeadingRegex.Match((headingMarkdown ?? string.Empty).Trim());
        if (!match.Success) {
            return false;
        }

        var level = match.Groups["level"].Value.Length;
        var text = match.Groups["text"].Value.Trim();
        if (text.Length == 0) {
            return false;
        }

        heading = new HeadingBlock(level, MarkdownReader.ParseInlineText(text, readerOptions));
        return true;
    }

    private static bool IsLegacyToolSlugHeading(HeadingBlock heading) {
        if (heading == null || heading.Level < 2 || heading.Level > 6) {
            return false;
        }

        return LegacyToolSlugHeadingRegex.IsMatch((heading.Text ?? string.Empty).Trim());
    }
}
