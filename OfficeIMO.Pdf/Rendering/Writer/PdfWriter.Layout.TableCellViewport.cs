namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void WriteTableCellViewportImage(PageImage image, TableCellContentFrame? clipFrame) {
            image.InlineDrawToken = AllocateInlineImageDrawToken(currentPage!);
            if (clipFrame.HasValue) {
                TableCellContentFrame clip = clipFrame.Value;
                new ContentStreamBuilder(sb).SaveState().Rectangle(clip.Left, clip.Top - clip.Height, clip.Width, clip.Height).ClipPath().EndPath();
            }
            sb.Append(image.InlineDrawToken);
            if (clipFrame.HasValue) new ContentStreamBuilder(sb).RestoreState();
        }
    }

    private readonly struct TableCellContentFrame {
        internal TableCellContentFrame(double left, double top, double width, double height) {
            Left = left;
            Top = top;
            Width = width;
            Height = height;
        }
        internal double Left { get; }
        internal double Top { get; }
        internal double Width { get; }
        internal double Height { get; }
    }

    private static double GetTableCellContentWidth(TableCellLayout cell, double cellWidth) {
        double width = cell.Viewport is { } viewport ? cellWidth * (viewport.ContentWidth / viewport.Width) : cellWidth;
        if (double.IsInfinity(width)) throw new ArgumentOutOfRangeException(nameof(cell), "The scaled cell text viewport width must remain finite.");
        return width;
    }

    private static TableCellContentFrame GetTableCellContentFrame(TableCellLayout cell, double left, double top, double width, double height) {
        if (cell.Viewport is not { } viewport) return new TableCellContentFrame(left, top, width, height);
        double fullWidth = width * (viewport.ContentWidth / viewport.Width);
        double fullHeight = height * (viewport.ContentHeight / viewport.Height);
        double fullLeft = left - width * (viewport.OffsetX / viewport.Width);
        double fullTop = top + height * (viewport.OffsetY / viewport.Height);
        if (double.IsInfinity(fullWidth) || double.IsInfinity(fullHeight) || double.IsInfinity(fullLeft) || double.IsInfinity(fullTop))
            throw new ArgumentOutOfRangeException(nameof(cell), "The scaled cell content viewport exceeds supported finite coordinates.");
        return new TableCellContentFrame(fullLeft, fullTop, fullWidth, fullHeight);
    }

    private static TableCellContentFrame? GetTableCellDiagonalFrame(TableCellLayout cell, double left, double bottom, double width, double height) =>
        cell.Viewport == null ? null : GetTableCellContentFrame(cell, left, bottom + height, width, height);

    private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<TableBlock, ViewportRowGroups> TableViewportRowGroups = new();

    private sealed class ViewportRowGroups {
        internal ViewportRowGroups(TableBlock table, int columns) {
            Ends = Enumerable.Repeat(-1, table.Rows.Count).ToArray();
            var anchorEnds = Enumerable.Repeat(-1, table.Rows.Count).ToArray();
            for (int row = 0; row < table.Rows.Count; row++) {
                foreach (TableCellLayout cell in GetTableCellLayouts(table, row, columns)) {
                    if (cell.Viewport != null) anchorEnds[row] = Math.Max(anchorEnds[row], row + cell.RowSpan - 1);
                }
            }
            for (int row = 0; row < anchorEnds.Length; row++) {
                if (anchorEnds[row] < row) continue;
                int start = row;
                int end = anchorEnds[row];
                while (row < end) { row++; end = Math.Max(end, anchorEnds[row]); }
                for (int member = start; member <= end; member++) Ends[member] = end;
            }
        }
        internal int[] Ends { get; }
    }

    private static int[] GetTableViewportRowGroups(TableBlock table, int columns) =>
        TableViewportRowGroups.GetValue(table, key => new ViewportRowGroups(key, columns)).Ends;

    private static bool TableRowHasViewport(TableBlock table, int row, int columns) => GetTableViewportRowGroups(table, columns)[row] >= row;

    private static bool StartsTableViewportRowGroup(int[] groups, int row) => groups[row] >= row && (row == 0 || groups[row - 1] < row);

    private static double GetTableViewportPlacementHeight(int[] groups, double[] rowHeights, int row, double rowGap) =>
        StartsTableViewportRowGroup(groups, row) ? GetTableRowsHeight(rowHeights, row, groups[row] - row + 1, rowGap) : rowHeights[row];

    private static void DrawTableCellDiagonals(System.Text.StringBuilder sb, PdfCellBorder border, double x, double y, double width, double height, TableCellContentFrame? frame, bool artifact) {
        if (!border.DiagonalUp && !border.DiagonalDown) return;
        if (frame.HasValue) {
            new ContentStreamBuilder(sb).SaveState().Rectangle(x, y, width, height).ClipPath().EndPath();
        }
        double left = frame?.Left ?? x;
        double bottom = frame.HasValue ? frame.Value.Top - frame.Value.Height : y;
        double right = left + (frame?.Width ?? width);
        double top = bottom + (frame?.Height ?? height);
        if (border.DiagonalUp) DrawCellDiagonalBorder(sb, ResolveCellBorderSide(border.DiagonalUpBorderSnapshot, border), left, bottom, right, top, diagonalUp: true, artifact);
        if (border.DiagonalDown) DrawCellDiagonalBorder(sb, ResolveCellBorderSide(border.DiagonalDownBorderSnapshot, border), left, bottom, right, top, diagonalUp: false, artifact);
        if (frame.HasValue) new ContentStreamBuilder(sb).RestoreState();
    }

    // Keep line slots and advances intact while omitting text wholly outside a
    // fragment. Partial intersections retain their original glyphs and PDF clip.
    private static void OmitInvisibleTableCellViewportLines(
        List<List<RichSeg>> lines, IReadOnlyList<double> heights,
        IReadOnlyList<PdfAlign?>? alignments, IReadOnlyList<double>? offsets, IReadOnlyList<double>? widths,
        PdfAlign alignment, double baseline, double textX, double textWidth,
        double clipX, double clipBottom, double clipWidth, double clipHeight, double leading, double fontSize, PdfOptions options, PdfStandardFont? baselineFont = null) {
        double lineY = baseline;
        for (int index = 0; index < lines.Count; index++) {
            List<RichSeg> line = lines[index];
            double inkBaseline = AdjustRichLineBaseline(lineY, line, options, fontSize, baselineFont);
            double lineWidth = MeasureRichLineWidth(line);
            double availableWidth = widths != null && index < widths.Count ? widths[index] : textWidth;
            double lineX = textX + (offsets != null && index < offsets.Count ? offsets[index] : 0D);
            PdfAlign lineAlignment = alignments != null && index < alignments.Count ? alignments[index] ?? alignment : alignment;
            if (lineAlignment == PdfAlign.Center) lineX += Math.Max(0D, (availableWidth - lineWidth) / 2D);
            else if (lineAlignment == PdfAlign.Right) lineX += Math.Max(0D, availableWidth - lineWidth);
            GetRichLineInkMetrics(line, options, out double ascender, out double descender);
            if (lineX + lineWidth <= clipX || lineX >= clipX + clipWidth ||
                inkBaseline + ascender <= clipBottom || inkBaseline - descender >= clipBottom + clipHeight)
                lines[index] = new List<RichSeg>();
            lineY -= index < heights.Count ? heights[index] : leading;
        }
    }
}
