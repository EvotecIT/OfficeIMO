namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private PdfPanelStyle ResolveContainerStyle(ContainerBlock container) {
            PdfPanelStyle style = container.UseDefaultPanelStyle ? currentOpts.DefaultPanelStyleSnapshot ?? container.Style : container.Style;
            if (style.FrameStartPaddingTopOverride.HasValue && Math.Abs(y - GetCurrentFramePageStartY()) <= .001D) {
                style = style.Clone();
                style.PaddingTopOverride = style.FrameStartPaddingTopOverride;
                style.TopBorder = style.FrameStartTopBorder;
            }
            return style;
        }

        private void RenderContainerBlock(
            ContainerBlock container,
            IPdfBlock? nextBlock,
            System.Collections.Generic.IList<IPdfBlock> blockList,
            int blockIndex) {
            PdfPanelStyle style = ResolveContainerStyle(container);
            double parentLeft = currentOpts.MarginLeft.ToPoints();
            double parentWidth = width;
            var frame = ResolveContainerFrame(container, style, parentLeft, parentWidth);
            double outerWidth = frame.Width;
            double outerX = frame.X;
            double contentWidth = frame.ContentWidth;

            double spacingBefore = ResolveTopLevelSpacingBefore(style.SpacingBefore.ToPoints());
            var firstVisual = MeasureContainerFirstVisual(container, style, outerX + style.PaddingX.ToPoints(), contentWidth, GetClosingContainerPadding());
            double firstVisualHeight = firstVisual.Height;
            double minimumStartHeight = spacingBefore + style.TopPadding + style.FragmentPaddingReservation + style.FragmentBottomInset.ToPoints() + firstVisualHeight;
            double firstContentHeight = minimumStartHeight - spacingBefore;
            RunningFrameAdmissionAttempt? runningAdmission = null;
            void RemeasureContainerStart() {
                style = ResolveContainerStyle(container);
                spacingBefore = ResolveTopLevelSpacingBefore(style.SpacingBefore.ToPoints());
                parentLeft = currentOpts.MarginLeft.ToPoints(); parentWidth = width;
                frame = ResolveContainerFrame(container, style, parentLeft, parentWidth);
                outerWidth = frame.Width; outerX = frame.X; contentWidth = frame.ContentWidth;
                firstVisual = MeasureContainerFirstVisual(container, style, frame.X + style.PaddingX.ToPoints(), frame.ContentWidth, GetClosingContainerPadding());
                firstVisualHeight = firstVisual.Height;
                minimumStartHeight = spacingBefore + style.TopPadding + style.FragmentPaddingReservation + style.FragmentBottomInset.ToPoints() + firstVisualHeight;
                firstContentHeight = minimumStartHeight - spacingBefore;
            }
            bool AdvanceContainerAdmission(double requiredHeight) =>
                TryAdvanceForRunningContent(ref runningAdmission, requiredHeight, preserve => {
                    NewBlockFrame(preserveEmptyPage: preserve);
                    RemeasureContainerStart();
                });
            while (firstContentHeight > GetMaximumBlockContinuationHeight() + .001D &&
                !CanContainerFitAnotherFlowWidth(container, style, parentWidth, wholeContainer: false)) {
                if (!AdvanceContainerAdmission(firstContentHeight))
                    throw new ArgumentException("Element padding and its first content cannot fit within the available page height.");
            }
            // Move an oversized kept paragraph before opening its decoration,
            // so an unused panel is not mistaken for an already painted fragment.
            while (firstContentHeight > GetMaximumBlockContinuationHeight() + .001D ||
                ShouldAdvanceForBlockHeight(minimumStartHeight) ||
                firstVisual.Overflow != PdfKeepTogetherOverflowBehavior.Reject &&
                (y < GetCurrentFramePageStartY() - .001D || IsPartialPhysicalColumnFrame())) {
                NewBlockFrame(forcePhysicalPage: firstVisual.Overflow != PdfKeepTogetherOverflowBehavior.Reject &&
                    (firstVisual.Overflow == PdfKeepTogetherOverflowBehavior.SplitAcrossPages || IsPartialPhysicalColumnFrame()));
                RemeasureContainerStart();
                while (firstContentHeight > GetMaximumBlockContinuationHeight() + .001D &&
                    !CanContainerFitAnotherFlowWidth(container, style, parentWidth, wholeContainer: false)) {
                    if (!AdvanceContainerAdmission(firstContentHeight))
                        throw new ArgumentException("Element padding and its first content cannot fit within the available page height.");
                }
            }

            void AdmitKeptContainerFrame() {
                if (!style.KeepTogether) return;
                double? keepHeight = MeasureWholeBlockHeight(container, parentLeft, parentWidth, currentOpts.DefaultFontSize.ToPoints());
                double? fullPageKeepHeight = MeasureWholeBlockAtFrameStart(container, parentLeft, parentWidth, currentOpts.DefaultFontSize.ToPoints());
                if (!keepHeight.HasValue || !fullPageKeepHeight.HasValue) {
                    throw new NotSupportedException("KeepTogether requires element content whose height can be determined before rendering. Remove KeepTogether or move dynamic, multi-column, deferred-table, table-of-contents, or explicit page-boundary content outside the element.");
                }

                while (fullPageKeepHeight.Value > GetMaximumBlockContinuationHeight() + .001D &&
                    !CanContainerFitAnotherFlowWidth(container, style, parentWidth, wholeContainer: true) &&
                    AdvanceContainerAdmission(fullPageKeepHeight.Value)) {
                    keepHeight = MeasureWholeBlockHeight(container, parentLeft, parentWidth, currentOpts.DefaultFontSize.ToPoints());
                    fullPageKeepHeight = MeasureWholeBlockAtFrameStart(container, parentLeft, parentWidth, currentOpts.DefaultFontSize.ToPoints());
                    if (!keepHeight.HasValue || !fullPageKeepHeight.HasValue)
                        throw new NotSupportedException("KeepTogether requires measurable element content after a page transition.");
                }
                double fullPageHeight = GetMaximumBlockContinuationHeight();
                // Preserve the initial width search through partial columns until the whole group can be placed.
                bool widthAdaptiveKeep = fullPageKeepHeight.Value > fullPageHeight + .001D;
                if (widthAdaptiveKeep &&
                    !CanContainerFitAnotherFlowWidth(container, style, parentWidth, wholeContainer: true)) {
                    throw new ArgumentException("Container height exceeds the available page content height while KeepTogether is enabled.");
                }

                while (keepHeight.Value > GetMaximumBlockContinuationHeight() + .001D || ShouldAdvanceForBlockHeight(keepHeight.Value)) {
                    NewBlockFrame();
                    style = ResolveContainerStyle(container);
                    spacingBefore = ResolveTopLevelSpacingBefore(style.SpacingBefore.ToPoints());
                    parentLeft = currentOpts.MarginLeft.ToPoints(); parentWidth = width;
                    keepHeight = MeasureWholeBlockAtFrameStart(container, parentLeft, parentWidth, currentOpts.DefaultFontSize.ToPoints());
                    if (!keepHeight.HasValue)
                        throw new ArgumentException("Container height exceeds the available page content height while KeepTogether is enabled.");
                    while (keepHeight.Value > GetMaximumBlockContinuationHeight() + .001D &&
                        (!widthAdaptiveKeep || !CanContainerFitAnotherFlowWidth(container, style, parentWidth, wholeContainer: true))) {
                        if (!AdvanceContainerAdmission(keepHeight.Value))
                            throw new ArgumentException("Container height exceeds the available page content height while KeepTogether is enabled.");
                        keepHeight = MeasureWholeBlockAtFrameStart(container, parentLeft, parentWidth, currentOpts.DefaultFontSize.ToPoints());
                        if (!keepHeight.HasValue)
                            throw new NotSupportedException("KeepTogether requires measurable element content after a page transition.");
                    }
                }
            }
            AdmitKeptContainerFrame();
            if (style.KeepWithNext && nextBlock != null) {
                double? elementHeight = MeasureWholeBlockHeight(container, parentLeft, parentWidth, currentOpts.DefaultFontSize.ToPoints());
                if (!elementHeight.HasValue) {
                    throw new NotSupportedException("KeepWithNext requires element content whose height can be determined before rendering. Remove KeepWithNext or move dynamic, multi-column, deferred-table, table-of-contents, canvas, or explicit page-boundary content outside the element.");
                }

                double nextHeight = MeasureCurrentFrameKeepNextHeight(blockList, blockIndex + 1, parentLeft, parentWidth, currentOpts.DefaultFontSize.ToPoints(), elementHeight.Value);
                double keepHeight = elementHeight.Value + nextHeight;
                double fullPageHeight = GetMaximumBlockContinuationHeight();
                while (nextHeight > 0.001D && keepHeight <= fullPageHeight + 0.001D && ShouldAdvanceForBlockHeight(keepHeight)) {
                    NewBlockFrame();
                    RemeasureContainerStart();
                    AdmitKeptContainerFrame();
                    elementHeight = MeasureWholeBlockAtFrameStart(container, parentLeft, parentWidth, currentOpts.DefaultFontSize.ToPoints());
                    if (!elementHeight.HasValue) break;
                    nextHeight = MeasureCurrentFrameKeepNextHeight(blockList, blockIndex + 1, parentLeft, parentWidth, currentOpts.DefaultFontSize.ToPoints(), elementHeight.Value);
                    keepHeight = elementHeight.Value + nextHeight;
                    fullPageHeight = GetMaximumBlockContinuationHeight();
                }
            }

            if (style.KeepTogether) {
                ReserveAtomicFootnoteGroup(new[] { container },
                    () => MeasureWholeBlockHeight(container, parentLeft, parentWidth, currentOpts.DefaultFontSize.ToPoints()) ?? 0D,
                    () => { NewBlockFrame(); RemeasureContainerStart(); AdmitKeptContainerFrame(); });
            }
            y -= spacingBefore;
            if (style.AnchoredCanvas is { } anchoredCanvas) RenderParagraphCanvas(anchoredCanvas, y);
            if (style.TopPadding > 0) RecordFlowPlacement(y);
            else ResolveFloatingBookmarks(y);
            PdfCreateOptions parentOptions = currentOpts;
            double parentYStart = yStart;
            PdfCreateOptions pageOptions = currentPage!.Options;
            parentWidth = width;
            frame = ResolveContainerFrame(container, style, currentOpts.MarginLeft.ToPoints(), parentWidth);
            outerX = frame.X;
            outerWidth = frame.Width;
            contentWidth = frame.ContentWidth;
            IReadOnlyList<IPdfBlock> contentBlocks = ResolveContainerContentBlocks(container, style,
                outerX + style.PaddingX.ToPoints(), contentWidth, currentOpts.DefaultFontSize.ToPoints());
            var nestedOptions = currentOpts.Clone();
            nestedOptions.MarginLeft = global::OfficeIMO.OfficeLength.Points(outerX + style.PaddingX.ToPoints());
            nestedOptions.MarginRight = global::OfficeIMO.OfficeLength.Points(nestedOptions.PageWidth.ToPoints() - (outerX + outerWidth - style.PaddingX.ToPoints()));
            nestedOptions.MarginBottom += style.FragmentBottomInset;
            // Imported fixed table frames may extend beyond the physical page.
            // Their derived horizontal margins are coordinates, while the page
            // options and the positive finite content frame remain validated.
            nestedOptions.Validate(allowDerivedHorizontalOverflow: true);

            var scope = new ContainerRenderScope(style, outerX, outerWidth, pageOptions, parentOptions, nestedOptions, container);
            activeContainerScopes.Add(scope);
            currentOpts = nestedOptions;
            width = contentWidth;
            BeginContainerFragment(scope);
            try {
                ProcessBlocks(contentBlocks, container);
                if (!style.RepeatFragmentDecoration && !style.AllowTrailingPaddingOverflow &&
                    style.BottomPadding > y - scope.ParentOptions.MarginBottom.ToPoints() + .001D)
                    throw new ArgumentException("Element closing padding cannot fit within the available page height.");
                double bottomPadding = style.AllowTrailingPaddingOverflow ? style.BottomPadding :
                    Math.Min(style.BottomPadding, Math.Max(0D, y - scope.ParentOptions.MarginBottom.ToPoints()));
                y -= bottomPadding;
                scope.ConsumedClosingPadding = bottomPadding;
                scope.MoveTrailingSeparatorToNextFrame = ShouldMoveJoinedContainerSeparator(scope, blockList, blockIndex);
                FinalizeContainerFragment(scope);
                // Nested margin options can resolve new fallback mappings as well
                // as glyphs. Transfer both to the page's font resource owner.
                scope.ParentOptions.MergeFontProgramUsageFrom(scope.NestedOptions);
            } finally {
                activeContainerScopes.RemoveAt(activeContainerScopes.Count - 1);
                currentOpts = scope.ParentOptions;
                width = currentOpts.PageWidth.ToPoints() - currentOpts.MarginLeft.ToPoints() - currentOpts.MarginRight.ToPoints();
                yStart = activeColumnFlow?.Top ?? parentYStart;
                if (currentPage != null) {
                    currentPage.Options = scope.PageOptions;
                }
            }

            if (style.SpacingAfter.ToPoints() > 0D) {
                ConsumeSpacer(GetContainerSpacingAfter(style, y - currentOpts.MarginBottom.ToPoints()));
            }
        }

        private static double GetContainerSpacingAfter(PdfPanelStyle style, double available) =>
            style.ClipSpacingAfterAtFrameEnd ? Math.Min(style.SpacingAfter.ToPoints(), Math.Max(0D, available)) : style.SpacingAfter.ToPoints();

        private static (double X, double Width, double ContentWidth) ResolveContainerFrame(PdfPanelStyle style, double parentLeft, double parentWidth) {
            double outerWidth = ResolveContainerOuterWidth(style, parentWidth);
            ValidatePanelStyle(style, outerWidth);
            double outerX = style.Align switch {
                PdfAlign.Center => parentLeft + (parentWidth - outerWidth) / 2D,
                PdfAlign.Right => parentLeft + parentWidth - outerWidth,
                _ => parentLeft
            };
            double contentWidth = outerWidth - 2D * style.PaddingX.ToPoints();
            if (contentWidth <= 0.001D) throw new ArgumentException("Container padding must leave positive content width.");
            return (outerX, outerWidth, contentWidth);
        }

        private void PrepareActiveContainerScopesForPageBreak(int firstIndex = 0, int? endIndex = null) {
            if (activeContainerScopes.Count == 0 || currentPage == null) {
                return;
            }

            for (int index = (endIndex ?? activeContainerScopes.Count) - 1; index >= firstIndex; index--) {
                ContainerRenderScope scope = activeContainerScopes[index];
                double bottomPadding = Math.Min(scope.Style.GetFragmentBottomPadding(continues: true), Math.Max(0D, y - currentOpts.MarginBottom.ToPoints()));
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
                var frame = ResolveContainerFrame(scope.Container, scope.Style, currentOpts.MarginLeft.ToPoints(), width);
                scope.OuterX = frame.X;
                scope.OuterWidth = frame.Width;
                // Only the frame changes during continuation. Reuse child options
                // and their accumulated font usage instead of copying assets per page.
                scope.NestedOptions.MarginLeft = global::OfficeIMO.OfficeLength.Points(scope.OuterX + scope.Style.PaddingX.ToPoints());
                scope.NestedOptions.MarginTop = currentOpts.MarginTop;
                scope.NestedOptions.MarginBottom = currentOpts.MarginBottom + scope.Style.FragmentBottomInset;
                scope.NestedOptions.MarginRight = global::OfficeIMO.OfficeLength.Points(currentOpts.PageWidth.ToPoints() -
                    (scope.OuterX + scope.OuterWidth - scope.Style.PaddingX.ToPoints()));
                currentOpts = scope.NestedOptions;
                width = currentOpts.PageWidth.ToPoints() - currentOpts.MarginLeft.ToPoints() - currentOpts.MarginRight.ToPoints();
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
            scope.TableContentBottom = null;
            if (scope.Container?.FrameTable is { } table) tableFrameScopes[table] = scope;
            return top - Math.Min(scope.Style.GetFragmentTopPadding(scope.IsContinuation), Math.Max(0D, top - currentOpts.MarginBottom.ToPoints()));
        }

        private void FinalizeContainerFragment(ContainerRenderScope scope) {
            FinalizeContainerFragment(scope, y);
        }

        private void FinalizeContainerFragment(ContainerRenderScope scope, double bottom, bool continues = false) {
            if (scope.Container?.FrameTable is { } table) tableFrameScopes.Remove(table);
            bottom = ResolveContainerTableFrameBottom(scope, bottom, continues);
            if (!continues && scope.MoveTrailingSeparatorToNextFrame) bottom += scope.ConsumedClosingPadding;
            double fragmentHeight = scope.FragmentTop - bottom;
            if (fragmentHeight <= 0.001D) {
                return;
            }

            var decoration = new StringBuilder();
            bool top = scope.Style.RepeatFragmentDecoration || !scope.IsContinuation;
            bool end = scope.Style.RepeatFragmentDecoration || !continues;
            if (scope.Style.Background.HasValue) {
                DrawRoundedRowFill(decoration, scope.Style.Background.Value, scope.OuterX, bottom, scope.OuterWidth, fragmentHeight, scope.Style.CornerRadius.ToPoints(), top, top, end, end, emitGeneratedStructure);
            }

            DrawPanelBorder(decoration, scope.Style, scope.OuterX, bottom, scope.OuterWidth, fragmentHeight, emitGeneratedStructure,
                top, end && !scope.MoveTrailingSeparatorToNextFrame);
            DrawContainerTableFrame(decoration, scope, bottom, fragmentHeight, top, end);
            if (decoration.Length > 0) {
                sb.Insert(scope.InsertionIndex, decoration.ToString());
                pageDirty = true;
            }
        }

        private sealed class ContainerRenderScope {
            public double? TableContentBottom;
            public double ConsumedClosingPadding { get; set; }
            public bool MoveTrailingSeparatorToNextFrame { get; set; }
            public ContainerRenderScope(PdfPanelStyle style, double outerX, double outerWidth, PdfCreateOptions pageOptions,
                PdfCreateOptions parentOptions, PdfCreateOptions nestedOptions, ContainerBlock? container = null) {
                Style = style;
                Container = container;
                OuterX = outerX;
                OuterWidth = outerWidth;
                PageOptions = pageOptions;
                ParentOptions = parentOptions;
                NestedOptions = nestedOptions;
            }

            public PdfPanelStyle Style { get; }
            public ContainerBlock? Container { get; }
            public double OuterX { get; set; }
            public double OuterWidth { get; set; }
            public PdfCreateOptions PageOptions { get; set; }
            public PdfCreateOptions ParentOptions { get; set; }
            public PdfCreateOptions NestedOptions { get; set; }
            public int InsertionIndex { get; set; }
            public double FragmentTop { get; set; }
            public bool IsContinuation { get; set; }
        }
    }
}
