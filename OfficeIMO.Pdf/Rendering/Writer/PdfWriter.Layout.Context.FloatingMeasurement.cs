namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Returns the children of a wrapper that adds no pagination constraint of its own.</summary>
        private static IReadOnlyList<IPdfBlock>? GetTransparentMeasurementChildren(IPdfBlock? block) => block switch {
            SemanticBlock semantic => semantic.Blocks,
            LayerBlock layer => layer.Blocks,
            FlowBlock flow when !flow.IsReplayable && !flow.Options.KeepTogether && flow.Options.ShowIf == null &&
                flow.Options.MinimumRemainingHeight.ToPoints() == 0 && flow.Options.OverflowBehavior == PdfFlowOverflowBehavior.Continue => flow.StaticBlocks,
            _ => null
        };

        private static IEnumerable<IPdfBlock> ExpandTransparentMeasurementBlocks(IEnumerable<IPdfBlock> blocks) {
            foreach (var block in blocks) {
                IReadOnlyList<IPdfBlock>? nested = GetTransparentMeasurementChildren(block);
                if (nested != null) {
                    foreach (var child in ExpandTransparentMeasurementBlocks(nested)) yield return child;
                } else yield return block;
            }
        }

        // Virtual reservations use the same page-space geometry as painting, but never
        // emit content or mutate capture/destination state. The caller restores them.
        private bool TryMeasureFloatingTable(TableBlock table, double frameWidth, double fontSize, out double bottom) {
            var style = table.Style ?? currentOpts.DefaultTableStyleSnapshot ?? PdfTableStyles.Light();
            bottom = y;
            if (style.Position is not { } position) return false;
            double height = MeasureTableBlockHeight(table, frameWidth, fontSize, firstVisualOnly: false);
            int count = GetTableColumnCount(table);
            if (count == 0 || table.Rows.Count == 0) return true;
            style = PreparePairedTableBorders(table, style);
            double tableWidth = ResolveTableColumnLayout(table, currentOpts, style, count, frameWidth,
                GetTableBodyFontSize(style, fontSize), style.HeaderRowCount, table.Rows.Count - style.FooterRowCount).Width;
            double flowY = y;
            double left = PositionTableX(position, tableWidth, table, style) + style.HorizontalOffset.ToPoints();
            y = PositionTableY(position, height);
            // Measure a header or footer float where painting puts it: moved up onto the page.
            if (isRunningContent) y = Math.Max(y, Math.Min(height, currentOpts.PageHeight.ToPoints()));
            AvoidFloatingTable(position, left, tableWidth, height);
            bottom = y - height;
            var outside = GetFloatingTableOutsideEdges(position, table, style);
            var distance = GetFloatingTableTextDistances(position);
            if (position.WrapsText) floatingTables.Add((currentPage, left - outside.Left - distance.Left,
                left + tableWidth + outside.Right + distance.Right, y + position.DistanceTop.ToPoints(),
                bottom - position.DistanceBottom.ToPoints(), position.AllowOverlap));
            y = flowY;
            return true;
        }

        /// <summary>
        /// A float placed at the flow position that keeps with the next block starts a new page
        /// when that block cannot begin beside it, so it stays with the text it belongs to.
        /// </summary>
        private void KeepFloatingTableWithNext(TableBlock table, IList<IPdfBlock> blocks, int blockIndex) {
            PdfTableStyle style = table.Style ?? currentOpts.DefaultTableStyleSnapshot ?? PdfTableStyles.Light();
            if (isRunningContent || !style.KeepWithNext || style.Position is not { VerticalAnchor: PdfTableAnchor.Flow }) return;
            double keepHeight = MeasureKeepWithNextChainHeight(blocks, blockIndex, currentOpts.MarginLeft.ToPoints(), width,
                currentOpts.DefaultFontSize.ToPoints(), 0D);
            if (keepHeight <= GetMaximumBlockContinuationHeight() + .001D && ShouldAdvanceForBlockHeight(keepHeight)) NewBlockFrame();
        }
    }
}
