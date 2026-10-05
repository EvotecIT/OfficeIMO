namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private double imageMeasurementReservedHeight;
        private double containerMeasurementTopPadding;

        // Preflight visits containers before their render scopes exist. Preserve the
        // same image content height while recursively measuring those containers.
        private T MeasureWithContainerPaddingReservation<T>(double paddingY, Func<T> measure) {
            double saved = imageMeasurementReservedHeight;
            double savedTopPadding = containerMeasurementTopPadding;
            imageMeasurementReservedHeight += paddingY * 2D;
            containerMeasurementTopPadding += paddingY;
            try {
                return measure();
            } finally {
                imageMeasurementReservedHeight = saved;
                containerMeasurementTopPadding = savedTopPadding;
            }
        }

        private double? MeasureBlockSequence(
            IReadOnlyList<IPdfBlock> blocks,
            double frameX,
            double frameWidth,
            double fontSize,
            double initialConsumedHeight = 0D,
            double? initialY = null) {
            double savedY = y;
            var savedFloats = floatingTables.ToArray();
            try {
                if (initialY.HasValue) {
                    y = initialY.Value;
                }

                y -= initialConsumedHeight;
                double startY = y;
                double deepest = y;
                foreach (IPdfBlock block in ExpandTransparentMeasurementBlocks(blocks)) {
                    if (block is BookmarkBlock) {
                        continue;
                    }

                    if (block is TableBlock floating && TryMeasureFloatingTable(floating, frameWidth, fontSize, out double floatingBottom)) {
                        deepest = Math.Min(deepest, floatingBottom);
                        continue;
                    }
                    double? height = block is RichParagraphBlock paragraph && floatingTables.Count > savedFloats.Length
                        ? MeasureFloatingParagraph(paragraph, frameX, frameWidth, fontSize)
                        : MeasureWholeBlockHeight(block, frameX, frameWidth, fontSize);
                    if (!height.HasValue) {
                        return null;
                    }

                    if (floatingTables.Count > savedFloats.Length && block is not RichParagraphBlock && block is not SpacerBlock &&
                        block is not FlowBlock && block is not SemanticBlock && block is not LayerBlock)
                        AvoidFloatingBlock(height.Value);
                    y -= height.Value;
                    deepest = Math.Min(deepest, y);
                }

                return startY - deepest;
            } finally {
                y = savedY;
                floatingTables.Clear(); floatingTables.AddRange(savedFloats);
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
                double? contentHeight = MeasureWithContainerPaddingReservation(style.FragmentPaddingReservation, () => MeasureBlockSequence(
                    container.Blocks,
                    frameX + style.PaddingX,
                    contentWidth,
                    fontSize,
                    spacingBefore + style.PaddingY));
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

                double flowStart = y;
                try {
                    double? measured = MeasureBlockSequence(flow.StaticBlocks, frameX, frameWidth, fontSize);
                    if (flow.Options.KeepTogether && measured.HasValue) {
                        while (HasFloatingTables) {
                            double previousY = y;
                            AvoidFloatingBlock(measured.Value);
                            if (y >= previousY - 0.001) break;
                            measured = MeasureBlockSequence(flow.StaticBlocks, frameX, frameWidth, fontSize);
                            if (!measured.HasValue) return null;
                        }
                    }
                    return measured.HasValue ? flowStart - y + measured.Value : null;
                } finally { y = flowStart; }
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
            for (int index = activeColumnFlow?.ContainerDepth ?? 0; index < activeContainerScopes.Count; index++) {
                ContainerRenderScope scope = activeContainerScopes[index];
                pageStart -= scope.Style.GetFragmentTopPadding(scope.IsContinuation);
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
