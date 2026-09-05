namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private sealed class ColumnGroup {
            public IReadOnlyList<ColItem> Children = null!;
            public PdfPanelStyle? Style;
            public FlowSemanticScope? Semantic;
            public PdfLayoutPositionCapture? Capture;
            public bool KeepTogether;
            public double XOffset;
            public double OuterWidth;
            public int ItemCount;
            public ContainerRenderScope? Decoration;
            public double FragmentTop;
        }

        private sealed class ColGroupStart : ColItem {
            public ColumnGroup Group = null!;
        }

        private sealed class ColGroupEnd : ColItem {
            public ColumnGroup Group = null!;
        }

        private bool TryAddColumnGroup(List<ColItem> items, IPdfBlock block, double columnWidth, double columnXOffset) {
            IReadOnlyList<IPdfBlock> blocks;
            var group = new ColumnGroup { XOffset = columnXOffset, OuterWidth = columnWidth };
            double childX = columnXOffset;
            double childWidth = columnWidth;
            switch (block) {
                case ContainerBlock container:
                    blocks = container.Blocks;
                    group.Style = ResolveContainerStyle(container);
                    group.KeepTogether = ResolveContainerStyle(container).KeepTogether;
                    var frame = ResolveContainerFrame(ResolveContainerStyle(container), columnXOffset, columnWidth);
                    group.XOffset = frame.X;
                    group.OuterWidth = frame.Width;
                    childX = frame.X + ResolveContainerStyle(container).PaddingX;
                    childWidth = frame.ContentWidth;
                    break;
                case SemanticBlock semantic:
                    blocks = semantic.Blocks;
                    group.Semantic = new FlowSemanticScope(semantic.Role, semantic.AlternativeText);
                    break;
                case FlowBlock flow when PdfFlowNestingRules.IsColumnFlowSupported(flow):
                    blocks = flow.StaticBlocks!;
                    group.KeepTogether = flow.Options.KeepTogether;
                    group.Capture = flow.Capture;
                    break;
                default:
                    return false;
            }

            List<ColItem> children = BuildColumnItems(blocks, childWidth, childX);
            group.Children = children;
            group.ItemCount = children.Count + 2;
            items.Add(new ColGroupStart { Group = group, ColumnXOffset = columnXOffset, ColumnWidth = columnWidth });
            items.AddRange(children);
            items.Add(new ColGroupEnd { Group = group, ColumnXOffset = columnXOffset, ColumnWidth = columnWidth });
            return true;
        }

        private bool TryBeginColumnGroup(
            ColumnGroup group, List<ColumnGroup> activeGroups, List<ColItem> items, int itemIndex,
            double columnX, ref double cursor, ref double remaining, ref double consumed) {
            double padding = group.Style?.PaddingY ?? 0D;
            double before = ResolveColumnSpacingBefore(group.Style?.SpacingBefore ?? 0D, consumed);
            double firstHeight = MeasureColumnGroupFirstVisualHeight(group);
            double fullHeight = MeasureColumnGroupHeight(group, consumed);
            double fullPageHeight = GetFullPageContentHeight() - activeGroups.Sum(active => (active.Style?.PaddingY ?? 0D) * 2D);
            double fullPageConsumed = activeGroups.Sum(active => active.Style?.PaddingY ?? 0D);
            if (group.KeepTogether && MeasureColumnGroupHeight(group, fullPageConsumed) > fullPageHeight + 0.001D) {
                throw new ArgumentException("Keep-together element or component content exceeds the available column height.");
            }

            double needed = group.KeepTogether ? fullHeight : before + padding * 2D + firstHeight;
            int nextItemIndex = itemIndex + group.ItemCount;
            if (group.Style?.KeepWithNext == true && nextItemIndex < items.Count) {
                double keepHeight = fullHeight + MeasureColKeepWithNextChainHeight(items, nextItemIndex);
                if (keepHeight <= fullPageHeight + 0.001D) needed = Math.Max(needed, keepHeight);
            }

            if (needed > remaining + 0.001D) {
                if (consumed <= 0.001D && Math.Abs(y - yStart) <= 0.001D) {
                    throw new ArgumentException("Element padding and its first content cannot fit within the available column height.");
                }
                return false;
            }

            ConsumeColumnSpace(before, ref cursor, ref remaining, ref consumed);
            if (group.Capture != null && initializedPositionCaptures.Add(group.Capture)) group.Capture.BeginLayoutPass();
            activeGroups.Add(group);
            BeginColumnGroupFragment(group, columnX, ref cursor, ref remaining, ref consumed);
            return true;
        }

        private void BeginColumnGroupFragment(ColumnGroup group, double columnX, ref double cursor, ref double remaining, ref double consumed) {
            group.FragmentTop = cursor;
            if (group.Semantic != null) flowSemanticScopes.Add(group.Semantic);
            if (group.Style != null) {
                group.Decoration = new ContainerRenderScope(group.Style, columnX + group.XOffset, group.OuterWidth, currentPage!.Options);
                double next = BeginContainerFragment(group.Decoration, cursor);
                ConsumeColumnSpace(cursor - next, ref cursor, ref remaining, ref consumed);
                // Reserve bottom padding so an inner item cannot occupy the decoration's closing space.
                remaining -= group.Style.PaddingY;
            }
        }

        private void EndColumnGroupFragment(ColumnGroup group, double columnX, ref double cursor, ref double remaining, ref double consumed) {
            if (group.Style != null) {
                remaining += group.Style.PaddingY;
                ConsumeColumnSpace(group.Style.PaddingY, ref cursor, ref remaining, ref consumed);
                FinalizeContainerFragment(group.Decoration!, cursor);
            }
            if (group.Capture != null && group.FragmentTop - cursor > 0.001D) {
                group.Capture.Add(new PdfLayoutRegion(pages.Count + 1, columnX + group.XOffset, cursor, group.OuterWidth, group.FragmentTop - cursor));
            }
            if (group.Semantic != null) flowSemanticScopes.RemoveAt(flowSemanticScopes.Count - 1);
        }

        private void ResumeColumnGroups(List<ColumnGroup> groups, double columnX, ref double cursor, ref double remaining, ref double consumed) {
            double padding = groups.Sum(group => (group.Style?.PaddingY ?? 0D) * 2D);
            if (padding > remaining + 0.001D) {
                throw new ArgumentException("Nested element padding exceeds the available column height.");
            }
            foreach (ColumnGroup group in groups) BeginColumnGroupFragment(group, columnX, ref cursor, ref remaining, ref consumed);
        }

        private void FinishColumnGroupsFragment(List<ColumnGroup> groups, double columnX, ref double cursor, ref double remaining, ref double consumed) {
            for (int index = groups.Count - 1; index >= 0; index--) {
                EndColumnGroupFragment(groups[index], columnX, ref cursor, ref remaining, ref consumed);
            }
        }

        private static void ConsumeColumnSpace(double height, ref double cursor, ref double remaining, ref double consumed) {
            cursor -= height;
            remaining -= height;
            consumed += height;
        }

        private double MeasureColumnGroupHeight(ColumnGroup group, double consumedBefore) {
            double before = ResolveColumnSpacingBefore(group.Style?.SpacingBefore ?? 0D, consumedBefore);
            double padding = group.Style?.PaddingY ?? 0D;
            return before + padding * 2D + MeasureColumnGroupContentHeight(group, consumedBefore + before + padding)
                + (group.Style?.SpacingAfter ?? 0D);
        }

        private double MeasureColumnGroupContentHeight(ColumnGroup group, double initialConsumed) {
            double total = 0D;
            foreach (ColItem child in group.Children) total += MeasureColItemFullHeight(child, initialConsumed + total);
            return total;
        }

        private double MeasureColumnGroupFirstVisualHeight(ColumnGroup group) {
            foreach (ColItem child in group.Children) {
                if (child is ColBookmark) continue;
                return MeasureColItemFirstVisualHeight(child);
            }
            return 0D;
        }
    }
}
