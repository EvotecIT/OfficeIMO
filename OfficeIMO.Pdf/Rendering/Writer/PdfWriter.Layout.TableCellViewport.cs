namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
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

    private static bool TableRowHasViewport(TableBlock table, int row, int columns) =>
        GetTableCellLayouts(table, row, columns).Any(cell => cell.Viewport != null);

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
        double clipX, double clipBottom, double clipWidth, double clipHeight, double leading, double fontSize, PdfOptions options) {
        double lineY = baseline;
        for (int index = 0; index < lines.Count; index++) {
            List<RichSeg> line = lines[index];
            double inkBaseline = AdjustRichLineBaseline(lineY, line, options, fontSize);
            double lineWidth = MeasureRichLineWidth(line);
            double availableWidth = widths != null && index < widths.Count ? widths[index] : textWidth;
            double lineX = textX + (offsets != null && index < offsets.Count ? offsets[index] : 0D);
            PdfAlign lineAlignment = alignments != null && index < alignments.Count ? alignments[index] ?? alignment : alignment;
            if (lineAlignment == PdfAlign.Center) lineX += Math.Max(0D, (availableWidth - lineWidth) / 2D);
            else if (lineAlignment == PdfAlign.Right) lineX += Math.Max(0D, availableWidth - lineWidth);
            double ascender = 0D, descender = 0D;
            foreach (RichSeg segment in line) {
                double rise = TextRiseForBaseline(segment.FontSize, segment.Baseline);
                double size = EffectiveRichFontSize(segment.FontSize, segment.Baseline);
                ascender = Math.Max(ascender, rise + GetAscenderForOptions(segment.Font, segment.NamedFont, size, options));
                descender = Math.Max(descender, GetDescenderForOptions(segment.Font, segment.NamedFont, size, options) - rise);
            }
            if (lineX + lineWidth <= clipX || lineX >= clipX + clipWidth ||
                inkBaseline + ascender <= clipBottom || inkBaseline - descender >= clipBottom + clipHeight)
                lines[index] = new List<RichSeg>();
            lineY -= index < heights.Count ? heights[index] : leading;
        }
    }
}
