namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Projects independent item cursors into common page bands after final item sizing.
    /// Source rows advance independently; repeated paint occupies destination space only.
    /// </summary>
    private bool TryProjectParallelTableBands(IReadOnlyList<HtmlRenderFlowBlock> blocks,
        ParallelTableProjectionContext context, out IReadOnlyList<HtmlRenderFlowBlock> projected) {
        projected = blocks;
        // A single item may itself contain parallel floats or neighboring text.
        // Qualify that shared source interval before a footer can consume it.
        var cursors = blocks.Select(block => new ParallelTableCursor(QualifySharedTableRepetition(block))).ToArray();
        double bandStart = 0D;
        double capacity = context.Boundary.RemainingHeight;
        int band = 0;
        while (cursors.Any(cursor => !cursor.Complete)) {
            CheckCancellation();
            ChargeLayoutOperations(cursors.Length, "parallel table page-band planning");
            int pageNumber = context.PageNumber + band;
            if (pageNumber > _options.MaxPageCount)
                throw new InvalidOperationException("HTML rendering exceeded the configured maximum page count.");
            HtmlCssPageGeometry geometry = _pageRules.ResolveGeometry(pageNumber, context.PageName, _options);
            if (!SamePageGeometry(context.Geometry, geometry)
                || Math.Abs(context.PageHeight - ResolvePageBodyContentHeight(pageNumber, geometry)) > 0.0001D) {
                _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.PagePseudoGeometryPending,
                    "Parallel table repetition retains its source layout because subsequent page geometry or reserved body space differs.",
                    HtmlDiagnosticSeverity.Warning, blocks[0].Source);
                return false;
            }
            bool anyProgress = false;
            foreach (ParallelTableCursor cursor in cursors) {
                CheckCancellation();
                if (cursor.Complete) continue;
                double before = cursor.Offset;
                double used = 0D;
                while (!cursor.Complete && used < capacity - 0.0001D) {
                    ChargeLayoutOperation("parallel table item fragment");
                    double available = capacity - used;
                    if (cursor.Offset <= 0.0001D && cursor.Block.AvoidBreakInside
                        && cursor.Block.Height <= context.PageHeight + 0.0001D
                        && cursor.Block.Height > available + 0.0001D) break;
                    PagedBlockFragment? fragment = PlanPagedFragment(cursor.Block, cursor.Offset,
                        available, context.PageHeight,
                        canForceFragment: used <= 0.0001D && capacity >= context.PageHeight - 0.0001D);
                    if (fragment == null || fragment.End <= cursor.Offset + 0.0001D) break;
                    if (fragment.SuppressedHeader || fragment.SuppressedFooter || fragment.Forced)
                        cursor.Losses.Add(fragment);
                    if (fragment.Header != null) {
                        AppendParallelTablePaint(cursor, fragment.Header.Visuals, bandStart + used);
                        used += fragment.HeaderHeight;
                    }
                    double sourceStart = cursor.Offset;
                    IReadOnlyList<HtmlRenderVisual> body = SliceBlockVisuals(cursor.Block, sourceStart, fragment.End,
                        ResolvePagedBackgroundWindow(sourceStart, fragment));
                    AppendParallelTablePaint(cursor, body, bandStart + used);
                    foreach (HtmlCssRunningStringAssignment assignment in cursor.Block.RunningStringAssignments) {
                        ChargeLayoutOperation("parallel table running-string projection");
                        if (assignment.Offset >= sourceStart - 0.0001D
                            && assignment.Offset < fragment.End - 0.0001D)
                            cursor.Assignments.Add(assignment.Translate(bandStart + used - sourceStart));
                    }
                    used += fragment.End - sourceStart;
                    cursor.Offset = fragment.End;
                    bool consumedFooter = false;
                    if (fragment.Footer != null) {
                        AppendParallelTablePaint(cursor, fragment.Footer.Visuals, bandStart + used);
                        used += fragment.FooterHeight;
                        if (cursor.Offset >= fragment.Footer.ContentEndsAt - 0.0001D) {
                            cursor.Offset = fragment.Footer.SourceEndsAt;
                            consumedFooter = true;
                        }
                    }
                    cursor.End = bandStart + used;
                    // Content after this item's final footer can use its remaining
                    // page space. No neighbor's source cursor is consumed here.
                    if (!consumedFooter) break;
                }
                anyProgress |= cursor.Offset > before + 0.0001D;
            }
            if (cursors.All(cursor => cursor.Complete)) break;
            foreach (ParallelTableCursor cursor in cursors.Where(cursor => !cursor.Complete))
                cursor.SourceBreaks.Add(cursor.Offset);
            if (!anyProgress && capacity >= context.PageHeight - 0.0001D) return false;
            // Keep the complete nonfinal band, including its blank tail, so a
            // root slice ends at the same boundary for every unfinished item.
            bandStart += capacity;
            foreach (ParallelTableCursor cursor in cursors.Where(cursor => !cursor.Complete))
                cursor.Breaks.Add(bandStart);
            capacity = context.PageHeight;
            band++;
        }
        foreach (ParallelTableCursor cursor in cursors) {
            foreach (PagedBlockFragment loss in cursor.Losses) ReportPagedFragmentLoss(cursor.Block, loss);
            foreach (double sourceBreak in cursor.SourceBreaks)
                ReportSharedTableRepetitionLoss(cursor.Block, sourceBreak);
        }
        projected = cursors.Select(cursor => new HtmlRenderFlowBlock(
            cursor.Block.Width, cursor.End, cursor.Visuals,
            cursor.Block.BreakBefore, cursor.Block.BreakAfter, cursor.Block.AvoidBreakInside, cursor.Block.Source,
            breakOffsets: cursor.Breaks,
            pageName: cursor.Block.PageName,
            stackingZIndex: cursor.Block.StackingZIndex,
            stackingSourceOrder: cursor.Block.StackingSourceOrder,
            ownerElement: cursor.Block.OwnerElement,
            runningStringAssignments: cursor.Assignments,
            layoutViewportWidth: cursor.Block.LayoutViewportWidth,
            layoutViewportHeight: cursor.Block.LayoutViewportHeight)).ToArray();
        return true;
    }

    private void AppendParallelTablePaint(ParallelTableCursor cursor,
        IReadOnlyList<HtmlRenderVisual> visuals, double offset) {
        foreach (HtmlRenderVisual visual in visuals) {
            foreach (HtmlRenderVisual _ in EnumeratePageFloatVisuals(new[] { visual })) {
                CheckCancellation();
                ChargeLayoutOperation("parallel table fragment paint");
            }
            cursor.Visuals.Add(visual.Translate(0D, offset, cursor.Visuals.Count));
        }
    }

    private sealed class ParallelTableCursor {
        internal ParallelTableCursor(HtmlRenderFlowBlock block) => Block = block;
        internal HtmlRenderFlowBlock Block { get; }
        internal double Offset { get; set; }
        internal double End { get; set; }
        internal bool Complete => Offset >= Block.Height - 0.0001D;
        internal List<HtmlRenderVisual> Visuals { get; } = new();
        internal List<double> Breaks { get; } = new();
        internal List<double> SourceBreaks { get; } = new();
        internal List<HtmlCssRunningStringAssignment> Assignments { get; } = new();
        internal List<PagedBlockFragment> Losses { get; } = new();
    }
}
