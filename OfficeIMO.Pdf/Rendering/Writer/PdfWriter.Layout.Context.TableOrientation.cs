using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Measures and paints quarter-turn cell content in logical coordinates while retaining the physical cell's borders and padding.</summary>
        private void RenderOrientedTableCellContent(TableCellLayout cell, PdfTableStyle style, int row, int column,
            TableCellContentFrame frame, double cellX, double cellTop, double cellWidth, double cellHeight,
            PdfStandardFont font, double fontSize, double leading, double runFontSizeScale, RichParagraphBlock paragraph,
            string structureType, int? markedContentId, bool includeCellObjects = true) {
            bool clockwise = cell.TextRotation < 0;
            double physicalLeft = GetTableCellPaddingLeft(style, row, column);
            double physicalRight = GetTableCellPaddingRight(style, row, column);
            double physicalTop = GetTableCellPaddingTop(style, row, column);
            double physicalBottom = GetTableCellPaddingBottom(style, row, column);
            double left = clockwise ? physicalTop : physicalBottom;
            double right = clockwise ? physicalBottom : physicalTop;
            double top = clockwise ? physicalRight : physicalLeft;
            double bottom = clockwise ? physicalLeft : physicalRight;
            double width = Math.Max(1D, frame.Height - left - right);
            double height = Math.Max(0D, frame.Width - top - bottom);
            var layout = CreateTableCellTextLayout(cell, width, font, fontSize, leading, currentOpts,
                runFontSizeScale, style.MinimumShrinkFontSize ?? 6D, style.AutoFitWidthUsesContentMinimum);
            // Exact line spacing may be smaller than the glyphs. Retain every
            // column whose ink intersects the physical cell, preserving advances
            // even when an intervening column is wholly outside the clip.
            int count = layout.LineCount;
            double textHeight = MeasureTableCellTextHeight(layout, 0, count, leading);
            double objectsHeight = MeasureTableCellObjectStackHeight(cell, width);
            double contentHeight = textHeight + (objectsHeight > 0D ? objectsHeight + (string.IsNullOrEmpty(cell.Text) ? 0D : TableCellCheckBoxGap) : 0D);
            double unused = Math.Max(0D, height - contentHeight);
            double offset = GetTableCellVerticalAlignment(style, row, column) switch {
                PdfCellVerticalAlign.Middle => unused / 2D,
                PdfCellVerticalAlign.Bottom => unused,
                _ => 0D
            };
            double firstBaseline = frame.Width - top - offset - layout.TopSpacing - GetAscenderForOptions(font, fontSize, currentOpts) + style.RowBaselineOffset;
            var lines = StripRichLineLinksWhenCellLinked(SliceTableCellLines(layout, 0, count), cell.LinkUri, cell.LinkDestinationName);
            var heights = SliceTableCellLineHeights(layout, 0, count, leading);
            var alignments = SliceTableCellLineAlignments(layout, 0, count);
            var offsets = SliceTableCellLineXOffsets(layout, 0, count);
            var widths = SliceTableCellLineWidths(layout, 0, count, width);
            double clipX = clockwise ? frame.Top - cellTop : cellTop - cellHeight - frame.Top + frame.Height;
            double clipBottom = clockwise ? cellX - frame.Left : frame.Left + frame.Width - cellX - cellWidth;
            OmitInvisibleTableCellViewportLines(lines, heights, alignments, offsets, widths,
                paragraph.Align, firstBaseline, left, width, clipX, clipBottom, cellHeight, cellWidth,
                leading, fontSize, currentOpts, font);
            OfficeTransform transform = clockwise
                ? new OfficeTransform(0D, -1D, 1D, 0D, frame.Left, frame.Top)
                : new OfficeTransform(0D, 1D, -1D, 0D, frame.Left + frame.Width, frame.Top - frame.Height);

            // Physical clipping happens outside the turn, including full-cell viewports.
            new ContentStreamBuilder(sb).SaveState().Rectangle(cellX, cellTop - cellHeight, cellWidth, cellHeight).ClipPath().EndPath();
            try {
                RenderOpaqueEffectGroupInline(transform, () => {
                    WriteClippedRichParagraph(sb, paragraph, lines, heights, currentOpts, firstBaseline, fontSize, leading,
                        currentPage!.Annotations, 0D, 0D, frame.Height, frame.Width, left, width,
                        structureType: structureType, markedContentId: markedContentId, structurePage: currentPage,
                        lineAlignments: alignments,
                        lineXOffsets: offsets,
                        lineWidths: widths, baselineFont: font);
                    if (includeCellObjects && objectsHeight > 0D)
                        RenderTableCellObjects(currentPage!, cell, MapCellAlignment(paragraph.Align), left, width,
                            frame.Width - top - offset - (string.IsNullOrEmpty(cell.Text) ? 0D : textHeight + TableCellCheckBoxGap),
                            pageImage => WriteTableCellViewportImage(pageImage, null));
                });
            } finally {
                new ContentStreamBuilder(sb).RestoreState();
            }
        }

        private static PdfColumnAlign MapCellAlignment(PdfAlign align) => align switch {
            PdfAlign.Center => PdfColumnAlign.Center,
            PdfAlign.Right => PdfColumnAlign.Right,
            _ => PdfColumnAlign.Left
        };
    }
}
