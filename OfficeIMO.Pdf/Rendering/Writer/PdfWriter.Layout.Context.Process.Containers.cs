namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private PdfPanelStyle ResolveContainerStyle(ContainerBlock container) =>
            container.UseDefaultPanelStyle ? currentOpts.DefaultPanelStyleSnapshot ?? container.Style : container.Style;

        private void RenderContainerBlock(
            ContainerBlock container,
            IPdfBlock? nextBlock,
            System.Collections.Generic.IList<IPdfBlock> blockList,
            int blockIndex) {
            PdfPanelStyle style = ResolveContainerStyle(container);
            double parentLeft = currentOpts.MarginLeft;
            double parentWidth = width;
            var frame = ResolveContainerFrame(style, parentLeft, parentWidth);
            double outerWidth = frame.Width;
            double outerX = frame.X;
            double contentWidth = frame.ContentWidth;

            double spacingBefore = ResolveTopLevelSpacingBefore(style.SpacingBefore);
            double firstVisualHeight = container.Blocks.Count == 0
                ? 0D
                : MeasureWithContainerPaddingReservation(style, () =>
                    MeasureNextBlockFirstVisualHeight(container.Blocks[0], outerX + style.PaddingX, contentWidth, currentOpts.DefaultFontSize, allowTableFragments: true));
            double minimumStartHeight = spacingBefore + style.PaddingY + style.FragmentPaddingReservation + style.FragmentBottomInset + firstVisualHeight;
            if (style.PaddingY + style.FragmentPaddingReservation + style.FragmentBottomInset + firstVisualHeight > GetMaximumBlockContinuationHeight() + 0.001D) {
                throw new ArgumentException("Element padding and its first content cannot fit within the available page height.");
            }
            while (ShouldAdvanceForBlockHeight(minimumStartHeight)) {
                NewBlockFrame();
                spacingBefore = ResolveTopLevelSpacingBefore(style.SpacingBefore);
                parentLeft = currentOpts.MarginLeft; parentWidth = width;
                frame = ResolveContainerFrame(style, parentLeft, parentWidth);
                firstVisualHeight = container.Blocks.Count == 0 ? 0D : MeasureWithContainerPaddingReservation(style, () =>
                    MeasureNextBlockFirstVisualHeight(container.Blocks[0], frame.X + style.PaddingX, frame.ContentWidth, currentOpts.DefaultFontSize, allowTableFragments: true));
                minimumStartHeight = spacingBefore + style.PaddingY + style.FragmentPaddingReservation + style.FragmentBottomInset + firstVisualHeight;
                if (style.PaddingY + style.FragmentPaddingReservation + style.FragmentBottomInset + firstVisualHeight > GetMaximumBlockContinuationHeight() + .001D)
                    throw new ArgumentException("Element padding and its first content cannot fit within the available page height.");
            }

            if (style.KeepTogether) {
                double? keepHeight = MeasureWholeBlockHeight(container, parentLeft, parentWidth, currentOpts.DefaultFontSize);
                double? fullPageKeepHeight = MeasureWholeBlockAtFrameStart(container, parentLeft, parentWidth, currentOpts.DefaultFontSize);
                if (!keepHeight.HasValue || !fullPageKeepHeight.HasValue) {
                    throw new NotSupportedException("KeepTogether requires element content whose height can be determined before rendering. Remove KeepTogether or move dynamic, multi-column, deferred-table, table-of-contents, or explicit page-boundary content outside the element.");
                }

                double fullPageHeight = GetMaximumBlockContinuationHeight();
                if (fullPageKeepHeight.Value > fullPageHeight + 0.001D) {
                    throw new ArgumentException("Container height exceeds the available page content height while KeepTogether is enabled.");
                }

                while (ShouldAdvanceForBlockHeight(keepHeight.Value)) {
                    NewBlockFrame();
                    spacingBefore = ResolveTopLevelSpacingBefore(style.SpacingBefore);
                    parentLeft = currentOpts.MarginLeft; parentWidth = width;
                    keepHeight = MeasureWholeBlockAtFrameStart(container, parentLeft, parentWidth, currentOpts.DefaultFontSize);
                    if (!keepHeight.HasValue || keepHeight.Value > GetMaximumBlockContinuationHeight() + .001D)
                        throw new ArgumentException("Container height exceeds the available page content height while KeepTogether is enabled.");
                }
            }
            if (style.KeepWithNext && nextBlock != null) {
                double? elementHeight = MeasureWholeBlockHeight(container, parentLeft, parentWidth, currentOpts.DefaultFontSize);
                if (!elementHeight.HasValue) {
                    throw new NotSupportedException("KeepWithNext requires element content whose height can be determined before rendering. Remove KeepWithNext or move dynamic, multi-column, deferred-table, table-of-contents, canvas, or explicit page-boundary content outside the element.");
                }

                double nextHeight = MeasureCurrentFrameKeepNextHeight(blockList, blockIndex + 1, parentLeft, parentWidth, currentOpts.DefaultFontSize, elementHeight.Value);
                double keepHeight = elementHeight.Value + nextHeight;
                double fullPageHeight = GetMaximumBlockContinuationHeight();
                while (nextHeight > 0.001D && keepHeight <= fullPageHeight + 0.001D && ShouldAdvanceForBlockHeight(keepHeight)) {
                    NewBlockFrame();
                    spacingBefore = ResolveTopLevelSpacingBefore(style.SpacingBefore);
                    parentLeft = currentOpts.MarginLeft; parentWidth = width;
                    elementHeight = MeasureWholeBlockAtFrameStart(container, parentLeft, parentWidth, currentOpts.DefaultFontSize);
                    if (!elementHeight.HasValue) break;
                    nextHeight = MeasureCurrentFrameKeepNextHeight(blockList, blockIndex + 1, parentLeft, parentWidth, currentOpts.DefaultFontSize, elementHeight.Value);
                    keepHeight = elementHeight.Value + nextHeight;
                    fullPageHeight = GetMaximumBlockContinuationHeight();
                    if (style.KeepTogether && elementHeight.Value > fullPageHeight + 0.001D)
                        throw new ArgumentException("Container height exceeds the available page content height while KeepTogether is enabled.");
                }
            }

            y -= spacingBefore;
            if (style.AnchoredCanvas is { } anchoredCanvas) RenderParagraphCanvas(anchoredCanvas, y);
            if (style.PaddingY > 0) RecordFlowPlacement(y);
            else ResolveFloatingBookmarks(y);
            PdfOptions parentOptions = currentOpts;
            double parentYStart = yStart;
            PdfOptions pageOptions = currentPage!.Options;
            parentWidth = width;
            frame = ResolveContainerFrame(style, currentOpts.MarginLeft, parentWidth);
            outerX = frame.X;
            outerWidth = frame.Width;
            contentWidth = frame.ContentWidth;
            IReadOnlyList<IPdfBlock> contentBlocks = ResolveContainerContentBlocks(container, style,
                outerX + style.PaddingX, contentWidth, currentOpts.DefaultFontSize);
            var nestedOptions = currentOpts.Clone();
            nestedOptions.MarginLeft = outerX + style.PaddingX;
            nestedOptions.MarginRight = nestedOptions.PageWidth - (outerX + outerWidth - style.PaddingX);
            nestedOptions.MarginBottom += style.FragmentBottomInset;
            nestedOptions.Validate();

            var scope = new ContainerRenderScope(style, outerX, outerWidth, pageOptions, parentOptions, nestedOptions);
            activeContainerScopes.Add(scope);
            currentOpts = nestedOptions;
            width = contentWidth;
            BeginContainerFragment(scope);
            try {
                ProcessBlocks(contentBlocks, container);
                if (!style.RepeatFragmentDecoration && style.PaddingY > y - scope.ParentOptions.MarginBottom + .001D)
                    throw new ArgumentException("Element closing padding cannot fit within the available page height.");
                double bottomPadding = Math.Min(style.PaddingY, Math.Max(0D, y - scope.ParentOptions.MarginBottom));
                y -= bottomPadding;
                FinalizeContainerFragment(scope);
                // Nested margin options can resolve new fallback mappings as well
                // as glyphs. Transfer both to the page's font resource owner.
                scope.ParentOptions.MergeFontProgramUsageFrom(scope.NestedOptions);
            } finally {
                activeContainerScopes.RemoveAt(activeContainerScopes.Count - 1);
                currentOpts = scope.ParentOptions;
                width = currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight;
                yStart = activeColumnFlow?.Top ?? parentYStart;
                if (currentPage != null) {
                    currentPage.Options = scope.PageOptions;
                }
            }

            if (style.SpacingAfter > 0D) {
                ConsumeSpacer(style.SpacingAfter);
            }
        }

        private static (double X, double Width, double ContentWidth) ResolveContainerFrame(PdfPanelStyle style, double parentLeft, double parentWidth) {
            double outerWidth = style.MaxWidth.HasValue ? Math.Min(parentWidth, style.MaxWidth.Value) : parentWidth;
            ValidatePanelStyle(style, outerWidth);
            double outerX = style.Align switch {
                PdfAlign.Center => parentLeft + (parentWidth - outerWidth) / 2D,
                PdfAlign.Right => parentLeft + parentWidth - outerWidth,
                _ => parentLeft
            };
            double contentWidth = outerWidth - 2D * style.PaddingX;
            if (contentWidth <= 0.001D) throw new ArgumentException("Container padding must leave positive content width.");
            return (outerX, outerWidth, contentWidth);
        }

        private void PrepareActiveContainerScopesForPageBreak(int firstIndex = 0, int? endIndex = null) {
            if (activeContainerScopes.Count == 0 || currentPage == null) {
                return;
            }

            for (int index = (endIndex ?? activeContainerScopes.Count) - 1; index >= firstIndex; index--) {
                ContainerRenderScope scope = activeContainerScopes[index];
                double bottomPadding = Math.Min(scope.Style.GetFragmentBottomPadding(continues: true), Math.Max(0D, y - currentOpts.MarginBottom));
                y -= bottomPadding;
                FinalizeContainerFragment(scope, y, continues: true);
                scope.ParentOptions.MergeFontProgramUsageFrom(scope.NestedOptions);
            }
        }

        private void ResumeActiveContainerScopesOnNewPage(int firstIndex = 0, int? endIndex = null) {
            if (activeContainerScopes.Count == 0 || currentPage == null) {
                return;
            }

            for (int index = firstIndex; index < (endIndex ?? activeContainerScopes.Count); index++) {
                ContainerRenderScope scope = activeContainerScopes[index];
                scope.PageOptions = currentPage.Options;
                scope.ParentOptions = currentOpts;
                var frame = ResolveContainerFrame(scope.Style, currentOpts.MarginLeft, width);
                scope.OuterX = frame.X;
                scope.OuterWidth = frame.Width;
                // Only the frame changes during continuation. Reuse child options
                // and their accumulated font usage instead of copying assets per page.
                scope.NestedOptions.MarginLeft = scope.OuterX + scope.Style.PaddingX;
                scope.NestedOptions.MarginTop = currentOpts.MarginTop;
                scope.NestedOptions.MarginBottom = currentOpts.MarginBottom + scope.Style.FragmentBottomInset;
                scope.NestedOptions.MarginRight = currentOpts.PageWidth -
                    (scope.OuterX + scope.OuterWidth - scope.Style.PaddingX);
                currentOpts = scope.NestedOptions;
                width = currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight;
                scope.IsContinuation = true;
                BeginContainerFragment(scope);
            }
        }

        private void BeginContainerFragment(ContainerRenderScope scope) {
            y = BeginContainerFragment(scope, y);
        }

        private double BeginContainerFragment(ContainerRenderScope scope, double top) {
            scope.InsertionIndex = sb.Length;
            scope.FragmentTop = top;
            return top - Math.Min(scope.Style.GetFragmentTopPadding(scope.IsContinuation), Math.Max(0D, top - currentOpts.MarginBottom));
        }

        private void FinalizeContainerFragment(ContainerRenderScope scope) {
            FinalizeContainerFragment(scope, y);
        }

        private void FinalizeContainerFragment(ContainerRenderScope scope, double bottom, bool continues = false) {
            double fragmentHeight = scope.FragmentTop - bottom;
            if (fragmentHeight <= 0.001D) {
                return;
            }

            var decoration = new StringBuilder();
            bool top = scope.Style.RepeatFragmentDecoration || !scope.IsContinuation;
            bool end = scope.Style.RepeatFragmentDecoration || !continues;
            if (scope.Style.Background.HasValue) {
                DrawRoundedRowFill(decoration, scope.Style.Background.Value, scope.OuterX, bottom, scope.OuterWidth, fragmentHeight, scope.Style.CornerRadius, top, top, end, end, emitGeneratedStructure);
            }

            DrawPanelBorder(decoration, scope.Style, scope.OuterX, bottom, scope.OuterWidth, fragmentHeight, emitGeneratedStructure, top, end);
            if (decoration.Length > 0) {
                sb.Insert(scope.InsertionIndex, decoration.ToString());
                pageDirty = true;
            }
        }

        private sealed class ContainerRenderScope {
            public ContainerRenderScope(PdfPanelStyle style, double outerX, double outerWidth, PdfOptions pageOptions,
                PdfOptions parentOptions, PdfOptions nestedOptions) {
                Style = style;
                OuterX = outerX;
                OuterWidth = outerWidth;
                PageOptions = pageOptions;
                ParentOptions = parentOptions;
                NestedOptions = nestedOptions;
            }

            public PdfPanelStyle Style { get; }
            public double OuterX { get; set; }
            public double OuterWidth { get; set; }
            public PdfOptions PageOptions { get; set; }
            public PdfOptions ParentOptions { get; set; }
            public PdfOptions NestedOptions { get; set; }
            public int InsertionIndex { get; set; }
            public double FragmentTop { get; set; }
            public bool IsContinuation { get; set; }
        }
    }
}
