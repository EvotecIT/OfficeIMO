using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    // A block inside a non-atomic inline interrupts the surrounding line flow. Keep
    // the block's pagination contract instead of disguising it as an inline image.
    private HtmlInlineLayout LayoutInterruptedInlineRuns(
        IReadOnlyList<HtmlInlineRun> runs,
        double width,
        HtmlRenderBoxStyle paragraphStyle,
        IElement? formattingContainer) {
        var blocks = new List<HtmlRenderFlowBlock>();
        var pending = new List<HtmlInlineRun>();
        var captures = new List<InlinePaintCapture?>();
        var blockRuns = new List<HtmlInlineRun?>();
        void Flush() {
            if (pending.Count == 0) return;
            var capture = new InlinePaintCapture(this);
            HtmlInlineLayout inline = LayoutInlineRuns(pending, width, blocks.Count == 0 ? paragraphStyle : WithoutTextIndent(paragraphStyle), formattingContainer, paintCapture: capture);
            pending.Clear();
            if (inline.Height <= 0D || inline.Visuals.Count == 0) return;
            captures.Add(capture);
            blockRuns.Add(null);
            blocks.Add(new HtmlRenderFlowBlock(width, inline.Height, inline.Visuals,
                HtmlPageBreakTarget.None, HtmlPageBreakTarget.None, false,
                formattingContainer == null ? "anonymous-block" : HtmlRenderStyleResolver.DescribeSource(formattingContainer),
                inline.BreakOffsets, inline.BreakOffsets, paragraphStyle.Orphans, paragraphStyle.Widows,
                pageName: paragraphStyle.PageName, runningStringAssignments: inline.RunningStringAssignments,
                layoutViewportWidth: ActiveSurfaceWidth, layoutViewportHeight: _activePageGeometry.Height));
        }
        foreach (HtmlInlineRun run in runs) {
            if (!run.IsBlockInterruption) {
                pending.Add(run);
                continue;
            }
            Flush();
            HtmlRenderFlowBlock block = run.AtomicBlock!;
            if (run.LinkUri != null) {
                var visuals = block.Visuals.ToList();
                OfficeShape link = OfficeShape.Rectangle(Math.Max(0.01D, block.Width), Math.Max(0.01D, block.Height));
                link.FillColor = null;
                link.StrokeWidth = 0D;
                visuals.Add(new HtmlRenderShape(link, 0D, 0D, visuals.Count, run.LinkUri, run.Source));
                block = block.WithVisuals(visuals);
            }
            block = block.TranslatePaint(run.PaintOffsetX, run.PaintOffsetY);
            blocks.Add(block);
            captures.Add(null);
            blockRuns.Add(run);
        }
        Flush();
        bool vertical = IsVerticalWritingMode(paragraphStyle.WritingMode);
        if (!vertical) CollapseInterruptedBlockMargins(blocks);
        var combined = new InlinePaintCapture(this);
        var breaks = new List<double>();
        var assignments = new List<HtmlCssRunningStringAssignment>();
        var forced = new List<HtmlRenderForcedBreak>();
        var lines = new List<HtmlRenderLineBreakGroup>();
        var continuations = new List<HtmlRenderContinuationGroup>();
        var trailing = new List<HtmlRenderTrailingGroup>();
        double height = 0D;
        bool rightToLeft = paragraphStyle.WritingMode == "vertical-rl" || paragraphStyle.WritingMode == "sideways-rl";
        double cursor = rightToLeft ? width : 0D;
        string? pageName = blocks.FirstOrDefault()?.PageName;
        for (int index = 0; index < blocks.Count; index++) {
            HtmlRenderFlowBlock block = blocks[index];
            InlinePaintCapture? capture = captures[index];
            if (capture == null) {
                capture = new InlinePaintCapture(this);
                HtmlInlineRun run = blockRuns[index]!;
                RecordInlineOwnerGeometry(run, formattingContainer, 0D, 0D, block.Width, block.Height,
                    capture.Bounds, decorationFragment: false);
                foreach (HtmlRenderVisual visual in block.Visuals) {
                    AddInlineOwnedVisual(capture.Visuals, capture.Owned, ApplyInlineElementSemantics(visual, run),
                        run.OwnerElement, formattingContainer);
                }
            }
            double offsetX = 0D;
            double offsetY = vertical ? 0D : height;
            if (vertical) {
                VerticalBlockBounds bounds = ResolveVerticalBlockBounds(block, paragraphStyle.LineHeight);
                offsetX = rightToLeft ? cursor - bounds.Right : cursor - bounds.Left;
                cursor = rightToLeft ? offsetX + bounds.Left : offsetX + bounds.Right;
            }
            combined.Merge(capture, offsetX, offsetY);
            if (index > 0 && !string.Equals(pageName, block.PageName, StringComparison.Ordinal)) {
                forced.Add(new HtmlRenderForcedBreak(offsetY, HtmlPageBreakTarget.Page, block.PageName, changesPageName: true));
            }
            pageName = block.PageName;
            if (block.BreakBefore != HtmlPageBreakTarget.None) forced.Add(new HtmlRenderForcedBreak(offsetY, block.BreakBefore));
            foreach (HtmlRenderForcedBreak item in block.ForcedBreaks) forced.Add(item.Translate(offsetY));
            foreach (double offset in block.BreakOffsets) breaks.Add(offsetY + offset);
            foreach (HtmlRenderLineBreakGroup group in block.LineBreakGroups) lines.Add(group.Translate(offsetY));
            foreach (HtmlRenderContinuationGroup group in block.ContinuationGroups) continuations.Add(group.Translate(offsetX, offsetY));
            foreach (HtmlRenderTrailingGroup group in block.TrailingGroups) trailing.Add(group.Translate(offsetX, offsetY));
            foreach (HtmlCssRunningStringAssignment assignment in block.RunningStringAssignments) assignments.Add(assignment.Translate(offsetY));
            height = vertical ? Math.Max(height, block.Height) : height + block.Height;
            if (block.BreakAfter != HtmlPageBreakTarget.None) forced.Add(new HtmlRenderForcedBreak(height, block.BreakAfter));
        }
        IReadOnlyList<HtmlRenderVisual> painted = ComposeInlinePositionedVisuals(
            combined.Visuals, combined.Owned, combined.Bounds, formattingContainer, isInlineContinuation: false);
        var aggregate = new HtmlRenderFlowBlock(width, height, painted,
            HtmlPageBreakTarget.None, HtmlPageBreakTarget.None, false,
            formattingContainer == null ? "anonymous-block" : HtmlRenderStyleResolver.DescribeSource(formattingContainer),
            breaks, lineBreakGroups: lines, continuationGroups: continuations, trailingGroups: trailing,
            hasCollapsibleMargins: blocks.Count > 0 && (blocks[0].HasCollapsibleMargins || blocks[blocks.Count - 1].HasCollapsibleMargins),
            collapsibleMarginTop: blocks.Count > 0 && blocks[0].HasCollapsibleMargins ? blocks[0].CollapsibleMarginTop : 0D,
            collapsibleMarginBottom: blocks.Count > 0 && blocks[blocks.Count - 1].HasCollapsibleMargins ? blocks[blocks.Count - 1].CollapsibleMarginBottom : 0D,
            ownerElement: blocks.FirstOrDefault()?.OwnerElement ?? blocks.LastOrDefault()?.OwnerElement,
            collapsesThrough: blocks.Count > 0 && blocks.All(block => block.CollapsesThrough),
            pageName: blocks.FirstOrDefault()?.PageName, runningStringAssignments: assignments, forcedBreaks: forced,
            layoutViewportWidth: ActiveSurfaceWidth, layoutViewportHeight: _activePageGeometry.Height);
        return new HtmlInlineLayout(painted, height, breaks, assignments, interruptedFlow: aggregate);
    }

    private static void CollapseInterruptedBlockMargins(List<HtmlRenderFlowBlock> blocks) {
        var adjoining = new AdjoiningMarginState();
        for (int index = 0; index < blocks.Count; index++) {
            HtmlRenderFlowBlock block = blocks[index];
            if (!block.HasCollapsibleMargins) {
                adjoining.Clear();
                continue;
            }
            if (adjoining.Count > 0) {
                adjoining.Add(block.CollapsibleMarginTop);
                if (!block.CollapsesThrough) block = block.AdjustLeadingFlowSpace(adjoining.Allocated - adjoining.Collapsed);
            }
            if (block.CollapsesThrough) {
                if (adjoining.Count == 0) adjoining.Reset(block.CollapsibleMarginTop);
                adjoining.Add(block.CollapsibleMarginBottom);
                block = block.AdjustTrailingFlowSpace(adjoining.Allocated - adjoining.Collapsed);
                adjoining.SetAllocated(adjoining.Collapsed);
            } else {
                adjoining.Reset(block.CollapsibleMarginBottom);
            }
            blocks[index] = block;
        }
    }

}
