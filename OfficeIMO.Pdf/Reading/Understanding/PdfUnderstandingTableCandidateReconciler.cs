namespace OfficeIMO.Pdf;

internal static class PdfUnderstandingTableCandidateReconciler {
    internal static IReadOnlyList<PdfUnderstandingTableCandidate> FilterAdditions(
        PdfLogicalPage page,
        IReadOnlyList<PdfUnderstandingTableCandidate> existing,
        IReadOnlyList<PdfUnderstandingTableCandidate> additions) {
        if (additions.Count == 0) return Array.Empty<PdfUnderstandingTableCandidate>();

        var accepted = new List<PdfUnderstandingTableCandidate>(additions.Count);
        for (int additionIndex = 0; additionIndex < additions.Count; additionIndex++) {
            PdfUnderstandingTableCandidate candidate = additions[additionIndex];
            bool duplicate = existing.Any(current => RepresentsSameVisualTable(page, current, candidate)) ||
                             accepted.Any(current => RepresentsSameVisualTable(page, current, candidate));
            if (!duplicate) accepted.Add(candidate);
        }
        return accepted.Count == 0 ? Array.Empty<PdfUnderstandingTableCandidate>() : accepted.AsReadOnly();
    }

    private static bool RepresentsSameVisualTable(
        PdfLogicalPage page,
        PdfUnderstandingTableCandidate left,
        PdfUnderstandingTableCandidate right) {
        if (!TryGetVisualBounds(page, left, out PdfVisualBounds leftBounds) ||
            !TryGetVisualBounds(page, right, out PdfVisualBounds rightBounds)) {
            return false;
        }

        double horizontalOverlap = Math.Max(0D, Math.Min(leftBounds.Right, rightBounds.Right) - Math.Max(leftBounds.Left, rightBounds.Left));
        double verticalOverlap = Math.Max(0D, Math.Min(leftBounds.Bottom, rightBounds.Bottom) - Math.Max(leftBounds.Top, rightBounds.Top));
        double narrowerWidth = Math.Min(leftBounds.Width, rightBounds.Width);
        double shorterHeight = Math.Min(leftBounds.Height, rightBounds.Height);
        if (narrowerWidth <= 0D || shorterHeight <= 0D) return false;

        double horizontalRatio = horizontalOverlap / narrowerWidth;
        double verticalRatio = verticalOverlap / shorterHeight;
        if (horizontalRatio >= 0.65D && verticalRatio >= 0.6D) return true;
        if (horizontalRatio < 0.5D || verticalRatio < 0.3D) return false;

        HashSet<string> leftCells = GetCellSignatures(left);
        HashSet<string> rightCells = GetCellSignatures(right);
        if (leftCells.Count == 0 || rightCells.Count == 0) return false;
        int shared = leftCells.Count <= rightCells.Count
            ? leftCells.Count(rightCells.Contains)
            : rightCells.Count(leftCells.Contains);
        return shared >= 2 && shared * 2 >= Math.Min(leftCells.Count, rightCells.Count);
    }

    private static bool TryGetVisualBounds(
        PdfLogicalPage page,
        PdfUnderstandingTableCandidate candidate,
        out PdfVisualBounds bounds) {
        if (candidate.Columns.Count == 0) {
            bounds = default;
            return false;
        }
        if (candidate.CoordinateSpace == PdfTableCoordinateSpace.VisualTopLeft) {
            PdfLogicalVisualBounds? visual = candidate.VisualBounds;
            double left = visual?.Left ?? candidate.Columns.Min(static column => Math.Min(column.From, column.To));
            double right = visual?.Right ?? candidate.Columns.Max(static column => Math.Max(column.From, column.To));
            double top = visual?.Top ?? Math.Min(candidate.YTop, candidate.YBottom);
            double bottom = visual?.Bottom ?? Math.Max(candidate.YTop, candidate.YBottom);
            bounds = new PdfVisualBounds(left, top, right, bottom);
        } else {
            double left = candidate.Columns.Min(static column => Math.Min(column.From, column.To));
            double right = candidate.Columns.Max(static column => Math.Max(column.From, column.To));
            double bottom = Math.Min(candidate.YBottom, candidate.YTop);
            double top = Math.Max(candidate.YBottom, candidate.YTop);
            bounds = page.TransformBoundsToVisual(left, bottom, right, top);
        }
        return bounds.Right > bounds.Left && bounds.Bottom > bounds.Top;
    }

    private static HashSet<string> GetCellSignatures(PdfUnderstandingTableCandidate candidate) {
        var result = new HashSet<string>(StringComparer.Ordinal);
        for (int rowIndex = 0; rowIndex < candidate.Rows.Count; rowIndex++) {
            IReadOnlyList<string> row = candidate.Rows[rowIndex];
            for (int columnIndex = 0; columnIndex < row.Count; columnIndex++) {
                string signature = NormalizeCell(row[columnIndex]);
                if (signature.Length > 0) result.Add(signature);
            }
        }
        return result;
    }

    private static string NormalizeCell(string value) {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        var result = new System.Text.StringBuilder(value.Length);
        bool pendingSpace = false;
        for (int index = 0; index < value.Length; index++) {
            char character = value[index];
            if (char.IsWhiteSpace(character)) {
                pendingSpace = result.Length > 0;
                continue;
            }
            if (pendingSpace) result.Append(' ');
            result.Append(char.ToUpperInvariant(character));
            pendingSpace = false;
        }
        return result.ToString();
    }
}
