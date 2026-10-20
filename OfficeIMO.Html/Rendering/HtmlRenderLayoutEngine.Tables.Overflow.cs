namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Applies the shared overflow owner after a cell's actual row/rowspan size
    /// and vertical alignment are known. Only its content is clipped; box paint
    /// and the table's semantic ownership keep their original geometry.
    /// </summary>
    private void ApplyTableCellContentOverflow(IReadOnlyList<TableRowLayout> rows, double spacing) {
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            foreach (TableCellLayout cell in rows[rowIndex].Cells) {
                if (cell.Style.OverflowX == "visible" && cell.Style.OverflowY == "visible") continue;
                CheckCancellation();
                HtmlInlineLayout inline = cell.Inline;
                ChargeLayoutOperations(1L + inline.Visuals.Count + inline.BreakOffsets.Count
                    + inline.BreakProgress.Count + inline.LineBreakOffsets.Count + inline.LineBreakGroups.Count
                    + inline.AvoidBreakRanges.Count + inline.SafeBreakRanges.Count + inline.ForcedBreaks.Count
                    + inline.RunningStringAssignments.Count + inline.FloatExclusions.Count
                    + inline.ContinuationGroups.Count + inline.TrailingGroups.Count,
                    "table used-cell overflow projection");
                double height = GetSpanningHeight(rows, rowIndex, cell.RowSpan, spacing);
                var content = new List<HtmlRenderVisual>();
                AppendOverflowContent(content, inline.Visuals, cell.Style, cell.Element,
                    -cell.Style.PaddingLeft,
                    cell.Style.BorderTopWidth - cell.ContentOffsetY,
                    Math.Max(0D, cell.Width - cell.Style.BorderLeftWidth - cell.Style.BorderRightWidth),
                    Math.Max(0D, height - cell.Style.BorderTopWidth - cell.Style.BorderBottomWidth),
                    propagateViewportOverflow: false);

                double paintExtent = inline.PagedPaintExtent;
                foreach (HtmlRenderVisual visual in content) {
                    // Use the geometry returned by the canonical clip owner,
                    // including overflow-clip-margin, rather than rebuilding it.
                    if (visual is HtmlRenderClipGroup { ClipVertical: true } clip)
                        paintExtent = Math.Min(paintExtent, Math.Max(0D, clip.ClipY + clip.ClipHeight));
                    else if (visual is HtmlRenderPathClipGroup pathClip)
                        paintExtent = Math.Min(paintExtent, Math.Max(0D, pathClip.ClipY + pathClip.ClipPath.Height));
                }
                cell.Inline = new HtmlInlineLayout(content, inline.Height, inline.BreakOffsets,
                    inline.RunningStringAssignments, inline.BreakProgress, inline.SupportsContinuationReflow,
                    inline.NormalFlowHeight, inline.LineBreakOffsets, inline.FloatExclusions,
                    interruptedFlow: inline.InterruptedFlow?.WithVisuals(content, pagedPaintExtent: paintExtent),
                    forcedBreaks: inline.ForcedBreaks, lineBreakGroups: inline.LineBreakGroups,
                    continuationGroups: inline.ContinuationGroups, trailingGroups: inline.TrailingGroups,
                    pagedPaintExtent: paintExtent, hasInFlowLineBoxes: inline.HasInFlowLineBoxes,
                    avoidBreakRanges: inline.AvoidBreakRanges, safeBreakRanges: inline.SafeBreakRanges);
            }
        }
    }
}
