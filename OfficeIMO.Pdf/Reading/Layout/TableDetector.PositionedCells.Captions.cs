namespace OfficeIMO.Pdf;

internal static partial class TableDetector {
    private static void AddWrappedPositionedCellGroups(
        List<PdfStructuredTable> result,
        List<(TextLayoutEngine.TextLine Line, PositionedRow Row)> anchors,
        List<TextLayoutEngine.TextLine> boundedLines,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        var group = new List<(TextLayoutEngine.TextLine Line, PositionedRow Row)>();
        TextLayoutEngine.TextLine? groupCaption = null;
        bool groupHasEmphasizedHeader = false;
        for (int index = 0; index < anchors.Count; index++) {
            cancellationCheck?.Invoke();
            (TextLayoutEngine.TextLine Line, PositionedRow Row) candidate = anchors[index];
            TextLayoutEngine.TextLine? separatingCaption = null;
            if (group.Count > 0) {
                bool aligned = PositionedRowsAlign(group[0].Row, candidate.Row);
                if (aligned && group.Count >= 3 &&
                    (groupCaption != null || groupHasEmphasizedHeader) &&
                    SamePositionedHeader(group[0].Row, candidate.Row)) {
                    TextLayoutEngine.TextLine? caption = FindPositionedTableCaption(candidate.Row, boundedLines,
                        consumeWork, cancellationCheck);
                    if (caption != null && caption.Y < group[group.Count - 1].Row.Y) separatingCaption = caption;
                }
                if (!aligned || group[group.Count - 1].Row.Y - candidate.Row.Y > 96D || separatingCaption != null) {
                    AddGroup(separatingCaption);
                    group.Clear();
                }
            }
            if (group.Count == 0) {
                groupHasEmphasizedHeader = HasEmphasizedText(candidate.Line);
                groupCaption = separatingCaption ?? (groupHasEmphasizedHeader
                    ? FindPositionedTableCaption(candidate.Row, boundedLines, consumeWork, cancellationCheck)
                    : null);
            }
            group.Add(candidate);
        }
        AddGroup(null);

        void AddGroup(TextLayoutEngine.TextLine? followingCaption) {
            if (groupCaption == null && followingCaption == null) {
                TryAddWrappedPositionedCellTable(result, group, boundedLines, consumeWork, cancellationCheck);
                return;
            }
            double top = groupCaption?.Y ?? double.PositiveInfinity;
            double bottom = followingCaption?.Y ?? double.NegativeInfinity;
            int start = FindPositionedLineAtOrBelow(boundedLines, top, consumeWork);
            while (start < boundedLines.Count && boundedLines[start].Y == top) {
                consumeWork?.Invoke(1);
                start++;
            }
            int end = FindPositionedLineAtOrBelow(boundedLines, bottom, consumeWork);
            consumeWork?.Invoke(end - start);
            List<TextLayoutEngine.TextLine> groupLines = boundedLines.GetRange(start, end - start);
            TryAddWrappedPositionedCellTable(result, group, groupLines, consumeWork, cancellationCheck,
                leadingCaption: groupCaption?.Text.Trim());
        }
    }

    private static bool SamePositionedHeader(PositionedRow first, PositionedRow next) {
        for (int index = 0; index < first.Cells.Count; index++) {
            string left = first.Cells[index].Text;
            string right = next.Cells[index].Text;
            if (string.Equals(left, right, StringComparison.Ordinal)) continue;
            if (!string.Equals(ContentStructureExtractor.NormalizeShattered(left).Trim(),
                    ContentStructureExtractor.NormalizeShattered(right).Trim(), StringComparison.Ordinal)) return false;
        }
        return LooksLikeHeaderRow(first.Cells.Select(static cell => cell.Text).ToArray());
    }

    // An outdented, smaller line isolated above the header belongs to the next table,
    // even when that table repeats the previous table's column labels and alignment.
    private static TextLayoutEngine.TextLine? FindPositionedTableCaption(
        PositionedRow header,
        List<TextLayoutEngine.TextLine> lines,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        double headerFont = header.Cells[0].SourceRuns.Max(static span => span.FontSize);
        double left = header.Cells[0].From;
        double right = header.Cells[header.Cells.Count - 1].To;
        double isolation = headerFont * 1.5D;
        TextLayoutEngine.TextLine? nearest = null;
        double precedingY = double.PositiveInfinity;
        int start = FindPositionedLineAtOrBelow(lines, header.Y + isolation * 3D, consumeWork);
        for (int index = start; index < lines.Count; index++) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            TextLayoutEngine.TextLine line = lines[index];
            if (line.Y <= header.Y + 1.5D) break;
            if (line.XStart > right || line.XEnd < left - headerFont || string.IsNullOrWhiteSpace(line.Text)) continue;
            if (nearest == null || line.Y < nearest.Y) {
                precedingY = nearest?.Y ?? double.PositiveInfinity;
                nearest = line;
            } else if (line.Y < precedingY) {
                precedingY = line.Y;
            }
        }
        if (nearest == null || nearest.Y - header.Y < isolation || nearest.Y - header.Y > isolation * 2D ||
            precedingY - nearest.Y < isolation || nearest.XStart > left - headerFont * 0.35D ||
            nearest.XEnd >= header.Cells[1].From ||
            nearest.Spans.Any(span => !string.IsNullOrWhiteSpace(span.Text) && span.FontSize >= headerFont)) return null;
        return nearest;
    }

    // The positioned pass receives lines from the existing bands in descending
    // baseline order. Search the neighboring region without scanning the page.
    private static int FindPositionedLineAtOrBelow(
        List<TextLayoutEngine.TextLine> lines, double y, Action<long>? consumeWork) {
        int low = 0;
        int high = lines.Count;
        while (low < high) {
            consumeWork?.Invoke(1);
            int middle = low + (high - low) / 2;
            if (lines[middle].Y > y) low = middle + 1;
            else high = middle;
        }
        return low;
    }
}
