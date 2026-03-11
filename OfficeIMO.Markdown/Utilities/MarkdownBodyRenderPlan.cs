namespace OfficeIMO.Markdown;

internal sealed class MarkdownBodyRenderPlan {
    private MarkdownBodyRenderPlan(
        IReadOnlyList<IMarkdownBlock> renderBlocks,
        TocPlaceholderBlock? sidebar,
        IReadOnlyList<FootnoteDefinitionBlock> footnotes) {
        RenderBlocks = renderBlocks;
        Sidebar = sidebar;
        Footnotes = footnotes;
    }

    internal IReadOnlyList<IMarkdownBlock> RenderBlocks { get; }
    internal TocPlaceholderBlock? Sidebar { get; }
    internal IReadOnlyList<FootnoteDefinitionBlock> Footnotes { get; }

    internal static MarkdownBodyRenderPlan Create(IReadOnlyList<IMarkdownBlock> blocks) {
        var renderBlocks = new List<IMarkdownBlock>();
        var footnotes = new List<FootnoteDefinitionBlock>();
        TocPlaceholderBlock? sidebar = null;

        for (int i = 0; i < blocks.Count; i++) {
            var block = blocks[i];

            if (block is FootnoteDefinitionBlock footnote) {
                footnotes.Add(footnote);
                continue;
            }

            if (ShouldSkipTocTitleHeading(blocks, i, block)) {
                continue;
            }

            if (block is TocPlaceholderBlock toc &&
                sidebar == null &&
                (toc.Options.Layout == TocLayout.SidebarLeft || toc.Options.Layout == TocLayout.SidebarRight)) {
                sidebar = toc;
                continue;
            }

            renderBlocks.Add(block);
        }

        return new MarkdownBodyRenderPlan(renderBlocks, sidebar, footnotes);
    }

    private static bool ShouldSkipTocTitleHeading(IReadOnlyList<IMarkdownBlock> blocks, int index, IMarkdownBlock block) {
        if (block is not HeadingBlock) {
            return false;
        }

        if (index + 1 >= blocks.Count || blocks[index + 1] is not TocPlaceholderBlock toc) {
            return false;
        }

        var options = toc.Options;
        return options.IncludeTitle &&
               (options.Layout == TocLayout.SidebarLeft ||
                options.Layout == TocLayout.SidebarRight ||
                options.Layout == TocLayout.Panel);
    }
}
