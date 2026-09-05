namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private double? MeasureBlockSequence(
            IReadOnlyList<IPdfBlock> blocks,
            double frameX,
            double frameWidth,
            double fontSize,
            double initialConsumedHeight = 0D,
            double? initialY = null) {
            double savedY = y;
            try {
                if (initialY.HasValue) {
                    y = initialY.Value;
                }

                y -= initialConsumedHeight;
                double total = 0D;
                for (int index = 0; index < blocks.Count; index++) {
                    IPdfBlock block = blocks[index];
                    if (block is BookmarkBlock) {
                        continue;
                    }

                    double? height = MeasureWholeBlockHeight(block, frameX, frameWidth, fontSize);
                    if (!height.HasValue) {
                        return null;
                    }

                    total += height.Value;
                    y -= height.Value;
                }

                return total;
            } finally {
                y = savedY;
            }
        }

        private double? MeasureWholeBlockHeight(IPdfBlock block, double frameX, double frameWidth, double fontSize) {
            if (block is SemanticBlock semantic) {
                return MeasureBlockSequence(semantic.Blocks, frameX, frameWidth, fontSize);
            }

            if (block is LayerBlock layer) {
                return MeasureBlockSequence(layer.Blocks, frameX, frameWidth, fontSize);
            }

            if (block is SectionBlock section) {
                if (section.Options.StartOnNewPage) {
                    return null;
                }

                var blocks = new List<IPdfBlock>(section.Blocks.Count + (section.Options.IncludeHeading ? 1 : 0));
                if (section.Options.IncludeHeading) {
                    blocks.Add(new HeadingBlock(
                        section.Options.Level,
                        section.Title,
                        PdfAlign.Left,
                        color: null,
                        style: section.Options.HeadingStyle));
                }

                blocks.AddRange(section.Blocks);
                return MeasureBlockSequence(blocks, frameX, frameWidth, fontSize);
            }

            if (block is ContainerBlock container) {
                PdfPanelStyle style = ResolveContainerStyle(container);
                double outerWidth = style.MaxWidth.HasValue ? Math.Min(frameWidth, style.MaxWidth.Value) : frameWidth;
                ValidatePanelStyle(style, outerWidth);
                double contentWidth = outerWidth - 2D * style.PaddingX;
                if (contentWidth <= 0.001D) {
                    throw new ArgumentException("Container padding must leave positive content width.");
                }

                double spacingBefore = ResolveTopLevelSpacingBefore(style.SpacingBefore);
                double? contentHeight = MeasureBlockSequence(
                    container.Blocks,
                    frameX + style.PaddingX,
                    contentWidth,
                    fontSize,
                    spacingBefore + style.PaddingY);
                return contentHeight.HasValue
                    ? spacingBefore + style.PaddingY + contentHeight.Value + style.PaddingY + style.SpacingAfter
                    : null;
            }

            if (block is FlowBlock flow) {
                if (flow.IsReplayable || flow.Options.ShowIf != null ||
                    flow.Options.MinimumRemainingHeight > 0D ||
                    flow.Options.OverflowBehavior != PdfFlowOverflowBehavior.Continue ||
                    flow.StaticBlocks == null) {
                    return null;
                }

                return MeasureBlockSequence(flow.StaticBlocks, frameX, frameWidth, fontSize);
            }

            if (block is PageBreakBlock or PageBlock or DeferredTableBlock or TableOfContentsBlock or
                MultiColumnBlock or ColumnBreakBlock or PdfCanvasBlock) {
                return null;
            }

            double height = MeasureKeepWithNextBlockHeight(block, frameX, frameWidth, fontSize);
            return height > 0D || block is SpacerBlock ? height : null;
        }

        private double GetCurrentFramePageStartY() {
            double pageStart = yStart;
            for (int index = 0; index < activeContainerScopes.Count; index++) {
                pageStart -= activeContainerScopes[index].Style.PaddingY;
            }

            return pageStart;
        }

        private double? MeasureWholeBlockAtFrameStart(IPdfBlock block, double frameX, double frameWidth, double fontSize) {
            double savedY = y;
            try {
                y = GetCurrentFramePageStartY();
                return MeasureWholeBlockHeight(block, frameX, frameWidth, fontSize);
            } finally {
                y = savedY;
            }
        }

        private double? MeasureBlockSequenceAtFrameStart(IReadOnlyList<IPdfBlock> blocks, double frameX, double frameWidth, double fontSize) {
            return MeasureBlockSequence(blocks, frameX, frameWidth, fontSize, initialY: GetCurrentFramePageStartY());
        }
    }
}
