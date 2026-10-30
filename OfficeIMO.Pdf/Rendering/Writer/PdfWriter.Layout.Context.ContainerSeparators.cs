namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        // A joined separator belongs to the next frame when its following element
        // cannot start here. Keep the flow cursor's reservation while omitting the
        // old frame's stroke and trailing decoration padding.
        private bool ShouldMoveJoinedContainerSeparator(ContainerRenderScope scope,
            IList<IPdfBlock> blocks, int index) {
            if (!scope.Style.BorderContinuesWithNext) return false;
            int nextIndex = index + 1;
            while (nextIndex < blocks.Count && IsNonVisualFlowMarker(blocks[nextIndex])) nextIndex++;
            if (nextIndex >= blocks.Count || blocks[nextIndex] is not ContainerBlock next) return false;
            PdfCreateOptions savedOptions = currentOpts;
            double savedY = y, savedWidth = width;
            activeContainerScopes.RemoveAt(activeContainerScopes.Count - 1);
            try {
                currentOpts = scope.ParentOptions;
                width = currentOpts.PageWidth.ToPoints() - currentOpts.MarginLeft.ToPoints() - currentOpts.MarginRight.ToPoints();
                PdfPanelStyle style = ResolveContainerStyle(next);
                if (!style.FrameStartPaddingTopOverride.HasValue) return false;
                double spacingAfter = GetContainerSpacingAfter(scope.Style, y - currentOpts.MarginBottom.ToPoints());
                if (spacingAfter > 0D && (spacingAfter > y - currentOpts.MarginBottom.ToPoints() || y - currentOpts.MarginBottom.ToPoints() <= .5D))
                    return true;
                y -= spacingAfter;
                var frame = ResolveContainerFrame(next, style, currentOpts.MarginLeft.ToPoints(), width);
                var visual = MeasureContainerFirstVisual(next, style, frame.X + style.PaddingX.ToPoints(), frame.ContentWidth);
                double needed = ResolveTopLevelSpacingBefore(style.SpacingBefore.ToPoints()) + style.TopPadding +
                    style.FragmentPaddingReservation + style.FragmentBottomInset.ToPoints() + visual.Height;
                if (ShouldAdvanceForBlockHeight(needed) || visual.Overflow != PdfKeepTogetherOverflowBehavior.Reject &&
                    (y < GetCurrentFramePageStartY() - .001D || IsPartialPhysicalColumnFrame())) return true;
                if (!style.KeepTogether && !style.KeepWithNext) return false;
                double? height = MeasureWholeBlockHeight(next, currentOpts.MarginLeft.ToPoints(), width, currentOpts.DefaultFontSize.ToPoints());
                if (!height.HasValue) return false;
                if (style.KeepTogether && ShouldAdvanceForBlockHeight(height.Value)) return true;
                if (style.KeepWithNext && nextIndex + 1 < blocks.Count) {
                    double nextHeight = MeasureCurrentFrameKeepNextHeight(blocks, nextIndex + 1, currentOpts.MarginLeft.ToPoints(),
                        width, currentOpts.DefaultFontSize.ToPoints(), height.Value);
                    double keepHeight = height.Value + nextHeight;
                    return nextHeight > .001D && keepHeight <= GetMaximumBlockContinuationHeight() + .001D &&
                        ShouldAdvanceForBlockHeight(keepHeight);
                }
                return false;
            } finally {
                currentOpts = savedOptions;
                y = savedY;
                width = savedWidth;
                activeContainerScopes.Add(scope);
            }
        }
    }
}
