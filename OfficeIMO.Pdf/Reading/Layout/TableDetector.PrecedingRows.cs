namespace OfficeIMO.Pdf;

internal static partial class TableDetector {
    /// <summary>
    /// Extends a table upwards with body rows the band detector lost because a neighbouring paragraph shares
    /// their band: a totals box beside a note keeps only its last rows otherwise ("VAT" and "Total" without
    /// "Subtotal"). A row is taken only when every cell is filled, it sits one row pitch above the table, nothing
    /// else on its baseline reaches into the table, and each cell starts or ends where that column's cells do.
    /// </summary>
    internal static void AttachPrecedingBodyRows(
        List<PdfStructuredTable> tables,
        IReadOnlyList<TextLayoutEngine.TextLine> pageLines,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        if (tables.Count == 0) return;
        var owned = new HashSet<TextLayoutEngine.TextLine>(tables.SelectMany(static table => table.SourceLines));
        foreach (PdfStructuredTable table in tables) {
            cancellationCheck?.Invoke();
            if (string.Equals(table.Kind, "leaders", StringComparison.Ordinal) || table.Columns.Count < 2 ||
                table.SourceLines.Count < 2 || table.Rows.Count != table.SourceLines.Count) continue;
            for (int attempt = 0; attempt < 8; attempt++) {
                cancellationCheck?.Invoke();
                if (!TryAttachPrecedingBodyRow(table, pageLines, owned, consumeWork)) break;
            }
        }
    }

    private static bool TryAttachPrecedingBodyRow(
        PdfStructuredTable table,
        IReadOnlyList<TextLayoutEngine.TextLine> pageLines,
        HashSet<TextLayoutEngine.TextLine> owned,
        Action<long>? consumeWork) {
        int columnCount = table.Columns.Count;
        double pitch = EstimateRowPitch(table, table);
        double left = table.Columns[0].From;
        double right = table.Columns[columnCount - 1].To;
        double firstRowY = table.SourceLines.Max(static line => line.Y);
        consumeWork?.Invoke(pageLines.Count);
        TextLayoutEngine.TextLine[] above = pageLines
            .Where(line => !owned.Contains(line) && line.Y > firstRowY + 0.5D && line.Y - firstRowY <= pitch * 1.35D &&
                line.XEnd > left && line.XStart < right && !string.IsNullOrWhiteSpace(line.Text))
            .ToArray();
        if (above.Length == 0) return false;
        double rowY = above.Min(static line => line.Y);
        TextLayoutEngine.TextLine[] row = above.Where(line => Math.Abs(line.Y - rowY) <= 1.5D).ToArray();
        if (row.Length != above.Length) return false;
        const double edgeTolerance = 3D;
        if (row.Any(line => line.XStart < left - edgeTolerance || line.XEnd > right + edgeTolerance)) return false;
        TextLayoutEngine.TextLine? candidate = row.Length == 1 ? row[0] : TryCombineSameBaselineLines(row.ToList());
        if (candidate is null || ContentStructureExtractor.TryParseListItemText(candidate.Text, out _, out _, out _)) return false;

        var splits = new List<double>(columnCount - 1);
        for (int column = 0; column + 1 < columnCount; column++) splits.Add(table.Columns[column].To);
        (double[] lefts, double[] rights) = GetColumnTextEdges(table.SourceLines, splits, columnCount);
        (double[] rowLefts, double[] rowRights, string[] cells) = GetRowCells(candidate, splits, columnCount);
        for (int column = 0; column < columnCount; column++) {
            if (string.IsNullOrWhiteSpace(cells[column]) || double.IsNaN(lefts[column])) return false;
            bool leftAligned = Math.Abs(rowLefts[column] - lefts[column]) <= edgeTolerance;
            bool rightAligned = Math.Abs(rowRights[column] - rights[column]) <= edgeTolerance;
            if (!leftAligned && !rightAligned) return false;
        }
        // Figures above figures: a header-like label row is left to the header attachment.
        bool sharesValueColumn = false;
        for (int column = 0; column < columnCount; column++) {
            if (!IsTabularValue(cells[column])) continue;
            if (table.Rows.Skip(1).All(existing => column < existing.Length && IsTabularValue(existing[column]))) sharesValueColumn = true;
        }
        if (!sharesValueColumn) return false;

        table.Rows.Insert(0, cells);
        var sourceLines = new List<TextLayoutEngine.TextLine>(table.SourceLines.Count + 1) { candidate };
        sourceLines.AddRange(table.SourceLines);
        table.SourceLines = sourceLines;
        table.SourceRuns = candidate.Spans.Concat(table.SourceRuns).Distinct().ToArray();
        table.YTop = Math.Max(table.YTop, candidate.Y);
        table.VisualBounds = null;
        foreach (TextLayoutEngine.TextLine line in row) owned.Add(line);
        return true;
    }

    private static (double[] Lefts, double[] Rights) GetColumnTextEdges(
        IEnumerable<TextLayoutEngine.TextLine> lines,
        List<double> splits,
        int columnCount) {
        var lefts = Enumerable.Repeat(double.NaN, columnCount).ToArray();
        var rights = Enumerable.Repeat(double.NaN, columnCount).ToArray();
        foreach (TextLayoutEngine.TextLine line in lines) {
            (double[] rowLefts, double[] rowRights, _) = GetRowCells(line, splits, columnCount);
            for (int column = 0; column < columnCount; column++) {
                if (double.IsNaN(rowLefts[column])) continue;
                lefts[column] = double.IsNaN(lefts[column]) ? rowLefts[column] : Math.Min(lefts[column], rowLefts[column]);
                rights[column] = double.IsNaN(rights[column]) ? rowRights[column] : Math.Max(rights[column], rowRights[column]);
            }
        }
        return (lefts, rights);
    }

    private static (double[] Lefts, double[] Rights, string[] Cells) GetRowCells(
        TextLayoutEngine.TextLine line,
        List<double> splits,
        int columnCount) {
        var lefts = Enumerable.Repeat(double.NaN, columnCount).ToArray();
        var rights = Enumerable.Repeat(double.NaN, columnCount).ToArray();
        foreach (PdfTextSpan span in line.Spans) {
            if (string.IsNullOrWhiteSpace(span.Text)) continue;
            int column = 0;
            while (column < splits.Count && span.X >= splits[column]) column++;
            double end = span.X + Math.Max(0D, span.Advance);
            lefts[column] = double.IsNaN(lefts[column]) ? span.X : Math.Min(lefts[column], span.X);
            rights[column] = double.IsNaN(rights[column]) ? end : Math.Max(rights[column], end);
        }
        return (lefts, rights, SplitBySplits(line, splits));
    }
}
