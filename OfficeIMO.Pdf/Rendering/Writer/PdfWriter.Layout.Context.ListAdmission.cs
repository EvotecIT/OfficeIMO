namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        // Preflight and placement share one terminal font-box/full-advance contract.
        private static double MeasurePreparedListAdmissionHeight(PreparedListLayout prepared, ContainerClosingSpace closingSpace) =>
            MeasurePreparedListHeight(prepared, frameFit: !closingSpace.RequiresFullAdvance) + closingSpace.Padding;

        private (double Height, PdfKeepTogetherOverflowBehavior Overflow) MeasureListFirstVisual(
            PdfListBlock list, double frameWidth, double fontSize, ContainerClosingSpace closingSpace) {
            PdfListStyle? listStyle = ResolveListStyle(list, currentOpts);
            double spacingBefore = listStyle?.SpacingBefore.ToPoints() ?? 0D;
            if (list.Items.Count == 0) return (spacingBefore, PdfKeepTogetherOverflowBehavior.Reject);

            double size = GetListFontSize(listStyle, fontSize);
            double markerSize = GetListMarkerFontSize(listStyle, size);
            double leading = Math.Max(GetListLeading(listStyle, size), GetListLeading(listStyle, markerSize));
            PreparedListLayout? prepared = null;
            PdfKeepTogetherOverflowBehavior overflow = PdfKeepTogetherOverflowBehavior.Reject;
            if (listStyle?.KeepTogether == true) {
                prepared = PrepareListLayout(list, frameWidth, fontSize, topLevelSpacing: true);
                double keptHeight = MeasurePreparedListAdmissionHeight(prepared, closingSpace);
                if (keptHeight <= GetMaximumBlockContinuationHeight() - imageMeasurementReservedHeight + .001D ||
                    listStyle.KeepTogetherOverflow == PdfKeepTogetherOverflowBehavior.Reject)
                    return (keptHeight, PdfKeepTogetherOverflowBehavior.Reject);
                // An oversized kept list yields to its splitting policy. Move it
                // before a decoration opens, as paragraph preflight does.
                overflow = listStyle.KeepTogetherOverflow;
            }

            if (closingSpace.RequiresFullAdvance ||
                listStyle?.LineSpacing?.FontLineBoxMultiplier != null && listStyle.LineSpacing.Rule == PdfLineSpacingRule.Multiple) {
                prepared ??= PrepareListLayout(list, frameWidth, fontSize, topLevelSpacing: true);
                TableCellTextLayout first = prepared.Items[0].TextLayout;
                return (spacingBefore + (first.LineCount == 0 ? 0D :
                    (prepared.Items.Count == 1 && first.LineCount == 1 ? closingSpace : default).GetTextHeight(
                        first.Lines[0], GetRichLineHeight(first.LineHeights, 0, leading), prepared.SpacingAfter)), overflow);
            }
            return (spacingBefore + leading, overflow);
        }
    }
}
