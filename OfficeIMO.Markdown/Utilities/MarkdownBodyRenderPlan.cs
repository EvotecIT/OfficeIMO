namespace OfficeIMO.Markdown;

internal sealed class MarkdownBodyRenderPlan {
    private MarkdownBodyRenderPlan(
        IReadOnlyList<IMarkdownBlock> renderBlocks,
        IBodySidebarMarkdownBlock? sidebar,
        IReadOnlyList<IFootnoteSectionMarkdownBlock> footnotes) {
        RenderBlocks = renderBlocks;
        Sidebar = sidebar;
        Footnotes = footnotes;
    }

    internal IReadOnlyList<IMarkdownBlock> RenderBlocks { get; }
    internal IBodySidebarMarkdownBlock? Sidebar { get; }
    internal IReadOnlyList<IFootnoteSectionMarkdownBlock> Footnotes { get; }

    internal static MarkdownBodyRenderPlan Create(IReadOnlyList<IMarkdownBlock> blocks) {
        var renderBlocks = new List<IMarkdownBlock>();
        var footnotes = new List<IFootnoteSectionMarkdownBlock>();
        IBodySidebarMarkdownBlock? sidebar = null;

        for (int i = 0; i < blocks.Count; i++) {
            var block = blocks[i];

            if (block is IFootnoteSectionMarkdownBlock footnote) {
                footnotes.Add(footnote);
                continue;
            }

            if (ShouldSkipTocTitleHeading(blocks, i, block)) {
                continue;
            }

            if (block is IBodySidebarMarkdownBlock toc &&
                sidebar == null &&
                toc.UsesSidebarLayout()) {
                sidebar = toc;
                continue;
            }

            renderBlocks.Add(block);
        }

        return new MarkdownBodyRenderPlan(renderBlocks, sidebar, footnotes);
    }

    private static bool ShouldSkipTocTitleHeading(IReadOnlyList<IMarkdownBlock> blocks, int index, IMarkdownBlock block) {
        if (block is not IHeadingMarkdownBlock) {
            return false;
        }

        if (index + 1 >= blocks.Count || blocks[index + 1] is not IBodySidebarMarkdownBlock toc) {
            return false;
        }

        return toc.SuppressesPrecedingHeadingTitle();
    }
}
