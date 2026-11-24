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
        // Most documents render every block in order; the filtered list and footnote list are created
        // only when a block is set aside.
        List<IMarkdownBlock>? renderBlocks = null;
        List<IFootnoteSectionMarkdownBlock>? footnotes = null;
        IBodySidebarMarkdownBlock? sidebar = null;

        for (int i = 0; i < blocks.Count; i++) {
            var block = blocks[i];
            bool setAside;

            if (block is IFootnoteSectionMarkdownBlock footnote) {
                (footnotes ??= new List<IFootnoteSectionMarkdownBlock>()).Add(footnote);
                setAside = true;
            } else if (ShouldSkipTocTitleHeading(blocks, i, block)) {
                setAside = true;
            } else if (block is IBodySidebarMarkdownBlock toc &&
                sidebar == null &&
                toc.UsesSidebarLayout()) {
                sidebar = toc;
                setAside = true;
            } else {
                setAside = false;
            }

            if (setAside) {
                if (renderBlocks == null) {
                    renderBlocks = new List<IMarkdownBlock>(blocks.Count);
                    for (int j = 0; j < i; j++) {
                        renderBlocks.Add(blocks[j]);
                    }
                }
            } else {
                renderBlocks?.Add(block);
            }
        }

        return new MarkdownBodyRenderPlan(
            renderBlocks ?? blocks,
            sidebar,
            (IReadOnlyList<IFootnoteSectionMarkdownBlock>?)footnotes ?? Array.Empty<IFootnoteSectionMarkdownBlock>());
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
