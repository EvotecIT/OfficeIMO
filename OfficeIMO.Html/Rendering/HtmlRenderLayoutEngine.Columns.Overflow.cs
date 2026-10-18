namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Reuses column fragmentation for a finite box's visible contents without
    /// consuming the space in which following normal-flow siblings are placed.
    /// </summary>
    private MultiColumnPlan ProjectMultiColumnPaintOverflow(MultiColumnPlan plan, double targetHeight,
        ColumnPageBands? pageBands) {
        var painted = new List<MultiColumnFragment>();
        var projected = new HashSet<HtmlRenderFlowBlock>();
        bool changed = false;
        foreach (MultiColumnFragment fragment in plan.Fragments) {
            CheckCancellation();
            ChargeLayoutOperation("column visible overflow projection");
            if (projected.Contains(fragment.Block)) continue;
            if (fragment.Start <= 0.0001D && fragment.Block.ColumnOverflowContent is { } overflow
                && overflow.Height > fragment.Block.Height + 0.0001D) {
                MultiColumnPlan paint = BuildMultiColumnPlan(new[] { overflow }, targetHeight,
                    _options.MaxColumnCount, throwOnLimit: true, notes: plan.Notes,
                    floats: plan.Floats, pageBands: pageBands,
                    initialColumn: fragment.Column, initialY: fragment.Y);
                bool ownsMetrics = true;
                foreach (ColumnOverflowPaintPiece piece in SplitColumnOverflowPaint(overflow.Visuals)) {
                    CheckCancellation();
                    ChargeLayoutOperation("column overflow paint selection");
                    long selectedNodeCount = CountColumnOverflowPaintNodes(piece.Visual);
                    if (piece.Positioned) {
                        // The containing box's positioned paint does not join
                        // the normal contents' independent column continuation.
                        ChargeLayoutOperations(selectedNodeCount, "column positioned paint materialization");
                        HtmlRenderFlowBlock retained = new(fragment.Block.Width, fragment.Block.Height,
                            new[] { piece.Visual }, HtmlPageBreakTarget.None, HtmlPageBreakTarget.None,
                            avoidBreakInside: false, source: fragment.Block.Source);
                        painted.Add(new MultiColumnFragment(retained, 0D, retained.Height,
                            fragment.Column, fragment.Y, carriesSourceFlow: false));
                    } else {
                        HtmlRenderFlowBlock selected = overflow.WithVisuals(new[] { piece.Visual });
                        foreach (MultiColumnFragment item in paint.Fragments) {
                            CheckCancellation();
                            ChargeLayoutOperations(selectedNodeCount, "column overflow selected materialization");
                            painted.Add(new MultiColumnFragment(selected, item.Start, item.End,
                                item.Column, item.Y, carriesSourceFlow: ownsMetrics,
                                backgroundWindowEnd: item.BackgroundWindowEnd));
                        }
                        ownsMetrics = false;
                    }
                }
                projected.Add(fragment.Block);
                changed = true;
            } else {
                painted.Add(fragment);
            }
        }
        return !changed ? plan : new MultiColumnPlan(plan.Fragments, plan.ColumnCount,
            plan.UsedHeight, plan.RuleBands, plan.BalanceExtent, painted, plan.Notes, plan.Floats);
    }

    /// <summary>Overlapping source paint permits a cut only when each crossing fragment permits it.</summary>
    private bool IsLegalColumnPaintCut(MultiColumnPlan plan, double offset) {
        if (!plan.HasPaintOverflow) return true;
        foreach (MultiColumnFragment fragment in plan.PaintFragments) {
            CheckCancellation();
            ChargeLayoutOperation("column overflow parallel cut");
            if (!fragment.CarriesSourceFlow) continue;
            if (offset <= fragment.Y + 0.0001D || offset >= fragment.Y + fragment.Height - 0.0001D) continue;
            double sourceOffset = fragment.Start + offset - fragment.Y;
            ChargeLayoutOperations((long)fragment.Block.BreakOffsets.Count + fragment.Block.SafeBreakRanges.Count,
                "column overflow source cut constraints");
            if (fragment.Block.BreakOffsets.Any(value => Math.Abs(value - sourceOffset) <= 0.0001D)
                || fragment.Block.SafeBreakRanges.Any(range => sourceOffset >= range.Start - 0.0001D
                    && sourceOffset <= range.End + 0.0001D)) continue;
            return false;
        }
        return true;
    }

    /// <summary>Retains paint order while separating positioned descendants from the contents that advance through columns.</summary>
    private IReadOnlyList<ColumnOverflowPaintPiece> SplitColumnOverflowPaint(IReadOnlyList<HtmlRenderVisual> visuals) {
        var result = new List<ColumnOverflowPaintPiece>();
        foreach (HtmlRenderVisual visual in visuals) result.AddRange(Split(visual));
        return result;

        IReadOnlyList<ColumnOverflowPaintPiece> Split(HtmlRenderVisual visual) {
            CheckCancellation();
            ChargeLayoutOperation("column overflow paint envelope");
            if (visual.IsOutOfFlowPaint || visual.PaintChildren is not { Count: > 0 } children)
                return new[] { new ColumnOverflowPaintPiece(visual,
                    visual.IsOutOfFlowPaint && !visual.FollowsColumnStaticPosition) };
            var pieces = new List<(bool Positioned, List<HtmlRenderVisual> Visuals)>();
            foreach (HtmlRenderVisual child in children) {
                foreach (ColumnOverflowPaintPiece piece in Split(child)) {
                    if (pieces.Count == 0 || pieces[pieces.Count - 1].Positioned != piece.Positioned)
                        pieces.Add((piece.Positioned, new List<HtmlRenderVisual>()));
                    pieces[pieces.Count - 1].Visuals.Add(piece.Visual);
                }
            }
            if (pieces.Count == 1)
                return new[] { new ColumnOverflowPaintPiece(visual, pieces[0].Positioned) };
            object identity = visual.PaintProjectionIdentity ?? new object();
            int textOwner = pieces.FindIndex(piece => !piece.Positioned);
            if (textOwner < 0) textOwner = 0;
            var projected = new List<ColumnOverflowPaintPiece>(pieces.Count);
            for (int index = 0; index < pieces.Count; index++) {
                CheckCancellation();
                ChargeLayoutOperation("column overflow split envelope materialization");
                (bool positioned, List<HtmlRenderVisual> content) = pieces[index];
                projected.Add(new ColumnOverflowPaintPiece(
                    visual.ProjectPaintChildren(content, 0D, 0D, visual.PaintOrder, ownsLogicalText: index == textOwner)
                        .IdentifyPaintProjection(identity), positioned));
            }
            return projected;
        }
    }

    private long CountColumnOverflowPaintNodes(HtmlRenderVisual visual) {
        long count = 0;
        var pending = new Stack<HtmlRenderVisual>();
        pending.Push(visual);
        while (pending.Count > 0) {
            CheckCancellation();
            ChargeLayoutOperation("column overflow source paint cost");
            HtmlRenderVisual current = pending.Pop();
            count++;
            if (current.PaintChildren is not { } children) continue;
            foreach (HtmlRenderVisual child in children) pending.Push(child);
        }
        return count;
    }

    /// <summary>A column-ending automatic static point belongs to the following column, while its box advances normally.</summary>
    private double ResolveColumnStaticPositionPaintExtent(IReadOnlyList<HtmlRenderVisual> visuals) {
        double extent = 0D;
        var pending = new Stack<HtmlRenderVisual>(visuals);
        while (pending.Count > 0) {
            CheckCancellation();
            ChargeLayoutOperation("column positioned static point extent");
            HtmlRenderVisual visual = pending.Pop();
            if (visual.IsOutOfFlowPaint) {
                if (visual.FollowsColumnStaticPosition)
                    extent = Math.Max(extent, visual.LayoutY + visual.LayoutHeight);
                continue;
            }
            if (visual.PaintChildren is not { } children) continue;
            foreach (HtmlRenderVisual child in children) pending.Push(child);
        }
        return extent;
    }

    private sealed class ColumnOverflowPaintPiece {
        internal ColumnOverflowPaintPiece(HtmlRenderVisual visual, bool positioned) {
            Visual = visual;
            Positioned = positioned;
        }
        internal HtmlRenderVisual Visual { get; }
        internal bool Positioned { get; }
    }
}
