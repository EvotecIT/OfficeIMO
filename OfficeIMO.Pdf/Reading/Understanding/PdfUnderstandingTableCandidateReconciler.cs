namespace OfficeIMO.Pdf;

internal static class PdfUnderstandingTableCandidateReconciler {
    internal static IReadOnlyList<PdfUnderstandingTableCandidate> Reconcile(
        PdfReadPage page,
        IReadOnlyList<PdfUnderstandingTableCandidate> existing,
        IReadOnlyList<PdfUnderstandingTableCandidate> additions) {
        return Reconcile(existing, additions, (left, right) => GetRelationship(page, left, right));
    }

    internal static IReadOnlyList<PdfUnderstandingTableCandidate> Reconcile(
        PdfLogicalPage page,
        IReadOnlyList<PdfUnderstandingTableCandidate> existing,
        IReadOnlyList<PdfUnderstandingTableCandidate> additions) {
        return Reconcile(existing, additions, (left, right) => GetRelationship(page, left, right));
    }

    private static IReadOnlyList<PdfUnderstandingTableCandidate> Reconcile(
        IReadOnlyList<PdfUnderstandingTableCandidate> existing,
        IReadOnlyList<PdfUnderstandingTableCandidate> additions,
        Func<PdfUnderstandingTableCandidate, PdfUnderstandingTableCandidate, CandidateRelationship> relationship) {
        var result = new List<PdfUnderstandingTableCandidate>(existing.Count + additions.Count);
        result.AddRange(existing);
        for (int additionIndex = 0; additionIndex < additions.Count; additionIndex++) {
            PdfUnderstandingTableCandidate candidate = additions[additionIndex];
            bool handled = false;
            for (int currentIndex = 0; currentIndex < result.Count; currentIndex++) {
                CandidateRelationship relation = relationship(result[currentIndex], candidate);
                if (relation == CandidateRelationship.AdditionRicher) {
                    result[currentIndex] = candidate;
                    handled = true;
                    break;
                }
                if (relation is CandidateRelationship.Duplicate or CandidateRelationship.ExistingRicher) {
                    handled = true;
                    break;
                }
            }
            if (!handled) result.Add(candidate);
        }
        return result.Count == 0 ? Array.Empty<PdfUnderstandingTableCandidate>() : result.AsReadOnly();
    }

    private static CandidateRelationship GetRelationship(
        PdfLogicalPage page,
        PdfUnderstandingTableCandidate left,
        PdfUnderstandingTableCandidate right) {
        if (!TryGetVisualBounds(page, left, out PdfVisualBounds leftBounds) ||
            !TryGetVisualBounds(page, right, out PdfVisualBounds rightBounds)) {
            return CandidateRelationship.Distinct;
        }
        return GetRelationship(left, right, leftBounds, rightBounds);
    }

    private static CandidateRelationship GetRelationship(
        PdfUnderstandingTableCandidate left,
        PdfUnderstandingTableCandidate right,
        PdfVisualBounds leftBounds,
        PdfVisualBounds rightBounds) {
        double horizontalOverlap = Math.Max(0D, Math.Min(leftBounds.Right, rightBounds.Right) - Math.Max(leftBounds.Left, rightBounds.Left));
        double verticalOverlap = Math.Max(0D, Math.Min(leftBounds.Bottom, rightBounds.Bottom) - Math.Max(leftBounds.Top, rightBounds.Top));
        double narrowerWidth = Math.Min(leftBounds.Width, rightBounds.Width);
        double shorterHeight = Math.Min(leftBounds.Height, rightBounds.Height);
        if (narrowerWidth <= 0D || shorterHeight <= 0D) return CandidateRelationship.Distinct;

        double horizontalRatio = horizontalOverlap / narrowerWidth;
        double verticalRatio = verticalOverlap / shorterHeight;
        if (horizontalRatio < 0.5D || verticalRatio < 0.3D) return CandidateRelationship.Distinct;

        Dictionary<string, int> leftRows = GetRowSignatures(left);
        Dictionary<string, int> rightRows = GetRowSignatures(right);
        if (HasStrictlyRicherContent(right, left) && ContainsAll(rightRows, leftRows)) {
            return CandidateRelationship.AdditionRicher;
        }
        if (HasStrictlyRicherContent(left, right) && ContainsAll(leftRows, rightRows)) {
            return CandidateRelationship.ExistingRicher;
        }
        int leftRowCount = CountSignatures(leftRows);
        int rightRowCount = CountSignatures(rightRows);
        if (leftRowCount == 0 || rightRowCount == 0) {
            return horizontalRatio >= 0.65D && verticalRatio >= 0.6D && left.Rows.Count == right.Rows.Count
                ? CandidateRelationship.Duplicate
                : CandidateRelationship.Distinct;
        }
        int shared = CountShared(leftRows, rightRows);
        return shared >= 1 && shared * 2 > Math.Min(leftRowCount, rightRowCount)
            ? CandidateRelationship.Duplicate
            : CandidateRelationship.Distinct;
    }

    private static CandidateRelationship GetRelationship(
        PdfReadPage page,
        PdfUnderstandingTableCandidate left,
        PdfUnderstandingTableCandidate right) {
        if (!TryGetVisualBounds(page, left, out PdfVisualBounds leftBounds) ||
            !TryGetVisualBounds(page, right, out PdfVisualBounds rightBounds)) {
            return CandidateRelationship.Distinct;
        }

        return GetRelationship(left, right, leftBounds, rightBounds);
    }

    private static bool ContainsAll(
        Dictionary<string, int> candidate,
        Dictionary<string, int> existing) =>
        existing.All(pair => candidate.TryGetValue(pair.Key, out int count) && count >= pair.Value);

    private static int CountSignatures(Dictionary<string, int> signatures) =>
        signatures.Values.Sum();

    private static int CountShared(
        Dictionary<string, int> left,
        Dictionary<string, int> right) {
        Dictionary<string, int> smaller = left.Count <= right.Count ? left : right;
        Dictionary<string, int> larger = ReferenceEquals(smaller, left) ? right : left;
        int shared = 0;
        foreach (KeyValuePair<string, int> pair in smaller) {
            if (larger.TryGetValue(pair.Key, out int count)) shared += Math.Min(pair.Value, count);
        }
        return shared;
    }

    private static bool HasStrictlyRicherContent(
        PdfUnderstandingTableCandidate candidate,
        PdfUnderstandingTableCandidate existing) {
        if (candidate.Rows.Count <= existing.Rows.Count) return false;
        return CountPopulatedCells(candidate) > CountPopulatedCells(existing);
    }

    private static int CountPopulatedCells(PdfUnderstandingTableCandidate candidate) {
        int count = 0;
        for (int rowIndex = 0; rowIndex < candidate.Rows.Count; rowIndex++) {
            IReadOnlyList<string> row = candidate.Rows[rowIndex];
            for (int columnIndex = 0; columnIndex < row.Count; columnIndex++) {
                if (!string.IsNullOrWhiteSpace(row[columnIndex])) count++;
            }
        }
        return count;
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

    private static bool TryGetVisualBounds(
        PdfReadPage page,
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

    private static Dictionary<string, int> GetRowSignatures(PdfUnderstandingTableCandidate candidate) {
        var result = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int rowIndex = 0; rowIndex < candidate.Rows.Count; rowIndex++) {
            IReadOnlyList<string> row = candidate.Rows[rowIndex];
            var signature = new System.Text.StringBuilder();
            for (int columnIndex = 0; columnIndex < row.Count; columnIndex++) {
                if (columnIndex > 0) signature.Append('\u001F');
                signature.Append(NormalizeCell(row[columnIndex]));
            }
            string value = signature.ToString();
            if (value.Trim('\u001F').Length == 0) continue;
            result[value] = result.TryGetValue(value, out int count) ? count + 1 : 1;
        }
        return result;
    }

    private static string NormalizeCell(string value) {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        string source = value.ToUpperInvariant();
        var result = new System.Text.StringBuilder(source.Length);
        bool pendingSpace = false;
        for (int index = 0; index < source.Length; index++) {
            char character = source[index];
            if (char.IsWhiteSpace(character)) {
                pendingSpace = result.Length > 0;
                continue;
            }
            if (pendingSpace) result.Append(' ');
            result.Append(character);
            pendingSpace = false;
        }
        return result.ToString();
    }

    private enum CandidateRelationship {
        Distinct,
        Duplicate,
        AdditionRicher,
        ExistingRicher
    }
}
