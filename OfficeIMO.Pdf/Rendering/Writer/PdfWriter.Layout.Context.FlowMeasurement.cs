namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private double imageMeasurementReservedHeight;
        private double containerMeasurementTopPadding;
        private double containerMeasurementBottomInset;

        // Preflight visits containers before their render scopes exist. Preserve the
        // same image content height while recursively measuring those containers.
        private T MeasureWithContainerPaddingReservation<T>(PdfPanelStyle style, Func<T> measure, bool isContinuation = false) {
            double saved = imageMeasurementReservedHeight;
            double savedTopPadding = containerMeasurementTopPadding;
            double savedBottomInset = containerMeasurementBottomInset;
            imageMeasurementReservedHeight += style.GetFragmentTopPadding(isContinuation) + Math.Max(style.BottomPadding, style.FragmentBottomInset);
            containerMeasurementTopPadding += style.GetFragmentTopPadding(isContinuation: true);
            containerMeasurementBottomInset += style.FragmentBottomInset;
            try {
                return measure();
            } finally {
                imageMeasurementReservedHeight = saved;
                containerMeasurementTopPadding = savedTopPadding;
                containerMeasurementBottomInset = savedBottomInset;
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
                var containerFrame = ResolveContainerFrame(container, style, frameX, frameWidth, fontSize);
                double contentWidth = containerFrame.ContentWidth;
                if (contentWidth <= 0.001D) {
                    throw new ArgumentException("Container padding must leave positive content width.");
                }

                double spacingBefore = ResolveTopLevelSpacingBefore(style.SpacingBefore);
                double? contentHeight = MeasureWithContainerPaddingReservation(style, () => MeasureBlockSequence(
                    container.Blocks,
                    containerFrame.X + style.PaddingX,
                    contentWidth,
                    fontSize,
                    spacingBefore + style.TopPadding));
                return contentHeight.HasValue
                    ? spacingBefore + style.TopPadding + Math.Max(contentHeight.Value, style.MinimumContentHeight) + style.BottomPadding + style.SpacingAfter
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
            return height > 0D || block is SpacerBlock or ShapeBlock ? height : null;
        }

        // Resolve against this frame after image scaling. The same padded block
        // sequence feeds page and row-column rendering, rather than relying on
        // a consumer's requested image dimensions.
        private IReadOnlyList<IPdfBlock> ResolveContainerContentBlocks(ContainerBlock container, PdfPanelStyle style,
            double frameX, double frameWidth, double fontSize, double reservedImageHeight = 0D) {
            if (style.MinimumContentHeight <= 0D) return container.Blocks;
            double savedReservation = imageMeasurementReservedHeight;
            double? height;
            try {
                imageMeasurementReservedHeight += reservedImageHeight;
                height = MeasureWithContainerPaddingReservation(style, () => MeasureBlockSequence(
                    container.Blocks, frameX, frameWidth, fontSize, style.TopPadding));
            } finally {
                imageMeasurementReservedHeight = savedReservation;
            }
            if (!height.HasValue) throw new NotSupportedException("Minimum panel content height requires measurable content.");
            double padding = style.MinimumContentHeight - height.Value;
            if (padding <= 0.001D) return container.Blocks;
            var result = new List<IPdfBlock>(container.Blocks.Count + 1);
            if (style.AlignContentToBottom) result.Add(new SpacerBlock(padding));
            result.AddRange(container.Blocks);
            if (!style.AlignContentToBottom) result.Add(new SpacerBlock(padding));
            return result;
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
