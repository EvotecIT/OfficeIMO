namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        // Keep preflight and placement on the same terminal font-box/full-advance contract.
        private static double MeasurePreparedListAdmissionHeight(PreparedListLayout prepared, ContainerClosingSpace closingSpace) =>
            MeasurePreparedListHeight(prepared, frameFit: !closingSpace.RequiresFullAdvance) + closingSpace.Padding;

        private (double Height, PdfKeepTogetherOverflowBehavior Overflow) MeasureListFirstVisual(
            PdfListBlock list, double frameWidth, double fontSize, ContainerClosingSpace closingSpace = default) {
            PdfListStyle? listStyle = ResolveListStyle(list, currentOpts);
            double size = GetListFontSize(listStyle, fontSize);
            double markerSize = GetListMarkerFontSize(listStyle, size);
            double leading = Math.Max(GetListLeading(listStyle, size), GetListLeading(listStyle, markerSize));
            string? firstItem = list.Items.Count > 0 ? list.Items[0] : null;
            if (firstItem == null)
                return (listStyle?.SpacingBefore ?? 0D, PdfKeepTogetherOverflowBehavior.Reject);

            PreparedListLayout? prepared = null;
            PdfKeepTogetherOverflowBehavior overflow = PdfKeepTogetherOverflowBehavior.Reject;
            if (listStyle?.KeepTogether == true) {
                prepared = PrepareListLayout(list, frameWidth, fontSize, topLevelSpacing: true);
                double keptHeight = MeasurePreparedListAdmissionHeight(prepared, closingSpace);
                if (keptHeight <= GetMaximumBlockContinuationHeight() + .001D ||
                    listStyle.KeepTogetherOverflow == PdfKeepTogetherOverflowBehavior.Reject)
                    return (keptHeight, PdfKeepTogetherOverflowBehavior.Reject);
                // Oversized imported keeps yield to their existing splitting policy.
                // Move before opening a decoration, as paragraph preflight does.
                overflow = listStyle.KeepTogetherOverflow;
            }

            if (closingSpace.RequiresFullAdvance ||
                listStyle?.LineSpacing?.FontLineBoxMultiplier != null && listStyle.LineSpacing.Rule == PdfLineSpacingRule.Multiple) {
                prepared ??= PrepareListLayout(list, frameWidth, fontSize, topLevelSpacing: true);
                TableCellTextLayout first = prepared.Items[0].TextLayout;
                double height = (listStyle?.SpacingBefore ?? 0D) + (first.LineCount == 0 ? 0D :
                    (prepared.Items.Count == 1 && first.LineCount == 1 ? closingSpace : default).GetTextHeight(
                        first.Lines[0], GetRichLineHeight(first.LineHeights, 0, leading), prepared.SpacingAfter));
                return (height, overflow);
            }
            return ((listStyle?.SpacingBefore ?? 0D) + leading, overflow);
        }
    }
}
