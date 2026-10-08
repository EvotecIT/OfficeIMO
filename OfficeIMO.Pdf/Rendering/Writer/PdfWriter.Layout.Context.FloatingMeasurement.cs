namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private static IEnumerable<IPdfBlock> ExpandTransparentMeasurementBlocks(IEnumerable<IPdfBlock> blocks) {
            foreach (var block in blocks) {
                IReadOnlyList<IPdfBlock>? nested = block switch {
                    SemanticBlock semantic => semantic.Blocks,
                    LayerBlock layer => layer.Blocks,
                    FlowBlock flow when !flow.IsReplayable && !flow.Options.KeepTogether && flow.Options.ShowIf == null &&
                        flow.Options.MinimumRemainingHeight == 0 && flow.Options.OverflowBehavior == PdfFlowOverflowBehavior.Continue => flow.StaticBlocks,
                    _ => null
                };
                if (nested != null) {
                    foreach (var child in ExpandTransparentMeasurementBlocks(nested)) yield return child;
                } else yield return block;
            }
        }

        // Virtual reservations use the same page-space geometry as painting, but never
        // emit content or mutate capture/destination state. The caller restores them.
        private bool TryMeasureFloatingTable(TableBlock table, double frameWidth, double fontSize, out double bottom) {
            var style = table.Style ?? currentOpts.DefaultTableStyleSnapshot ?? TableStyles.Light();
            bottom = y;
            if (style.Position is not { } position) return false;
            double height = MeasureTableBlockHeight(table, frameWidth, fontSize, firstVisualOnly: false);
            int count = GetTableColumnCount(table);
            if (count == 0 || table.Rows.Count == 0) return true;
            style = PreparePairedTableBorders(table, style);
            double tableWidth = ResolveTableColumnLayout(table, currentOpts, style, count, frameWidth,
                GetTableBodyFontSize(style, fontSize), style.HeaderRowCount, table.Rows.Count - style.FooterRowCount).Width;
            double flowY = y;
            double left = PositionTableX(position, tableWidth);
            y = PositionTableY(position, height);
            AvoidFloatingTable(position, left, tableWidth, height);
            bottom = y - height;
            floatingTables.Add((currentPage, left - position.DistanceLeft,
                left + tableWidth + position.DistanceRight, y + position.DistanceTop,
                bottom - position.DistanceBottom, position.AllowOverlap));
            y = flowY;
            return true;
        }

        private double MeasureFloatingParagraph(RichParagraphBlock paragraph, double frameX, double frameWidth, double fontSize, bool firstVisualOnly = false) {
            var style = EffectiveParagraphStyle(paragraph);
            fontSize = style?.FontSize ?? currentOpts.DefaultFontSize;
            double leading = GetParagraphLeading(style, fontSize);
            double spacing = ResolveTopLevelSpacingBefore(GetParagraphSpacingBefore(style));
            var textFrame = GetParagraphTextFrame(style, frameX, frameWidth);
            double start = y - spacing;
            var wrapped = WrapRichRunsCoreWithFirstLineOrigin(paragraph.Runs, textFrame.Width, fontSize,
                ChooseNormal(currentOpts.DefaultFont), leading, textFrame.FirstLineWidth,
                textFrame.FirstLineX - textFrame.X, GetParagraphTabStopWidth(style), currentOpts,
                style?.TabStops.ToArray(), (index, completedHeight, requiredHeight, minimumWidth) => {
                    double left = index == 0 ? textFrame.FirstLineX : textFrame.X;
                    double available = index == 0 ? textFrame.FirstLineWidth : textFrame.Width;
                    var frame = GetFloatingTextFrame(left, available, start - completedHeight, requiredHeight, minimumWidth);
                    return (frame.Width, frame.X - textFrame.X, frame.Gap);
                }, lineSpacing: style?.LineSpacing);
            return spacing + (firstVisualOnly ? wrapped.LineHeights.FirstOrDefault() : wrapped.LineHeights.Sum() + GetParagraphSpacingAfter(style, leading));
        }
    }
}
