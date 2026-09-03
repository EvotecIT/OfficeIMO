namespace OfficeIMO.Pdf;

internal static class PdfUnderstandingTableCandidateReconciler {
    internal static IReadOnlyList<PdfUnderstandingTableCandidate> Reconcile(
        PdfReadPage page,
        IReadOnlyList<PdfUnderstandingTableCandidate> existing,
        IReadOnlyList<PdfUnderstandingTableCandidate> additions,
        Action<long>? consumeWork = null,
        Action? cancellationCheck = null) {
        return Reconcile(
            existing,
            additions,
            candidate => CreateState(page, candidate, consumeWork, cancellationCheck),
            consumeWork,
            cancellationCheck);
    }

    internal static IReadOnlyList<PdfUnderstandingTableCandidate> Reconcile(
        PdfLogicalPage page,
        IReadOnlyList<PdfUnderstandingTableCandidate> existing,
        IReadOnlyList<PdfUnderstandingTableCandidate> additions,
        Action<long>? consumeWork = null,
        Action? cancellationCheck = null) {
        return Reconcile(
            existing,
            additions,
            candidate => CreateState(page, candidate, consumeWork, cancellationCheck),
            consumeWork,
            cancellationCheck);
    }

    private static IReadOnlyList<PdfUnderstandingTableCandidate> Reconcile(
        IReadOnlyList<PdfUnderstandingTableCandidate> existing,
        IReadOnlyList<PdfUnderstandingTableCandidate> additions,
        Func<PdfUnderstandingTableCandidate, CandidateState> createState,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        var result = new List<CandidateState>(existing.Count + additions.Count);
        for (int existingIndex = 0; existingIndex < existing.Count; existingIndex++) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            result.Add(createState(existing[existingIndex]));
        }
        for (int additionIndex = 0; additionIndex < additions.Count; additionIndex++) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            CandidateState candidate = createState(additions[additionIndex]);
            bool handled = false;
            for (int currentIndex = 0; currentIndex < result.Count; currentIndex++) {
                cancellationCheck?.Invoke();
                consumeWork?.Invoke(1);
                CandidateRelationship relation = GetRelationship(
                    result[currentIndex],
                    candidate,
                    consumeWork,
                    cancellationCheck);
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

        if (result.Count == 0) return Array.Empty<PdfUnderstandingTableCandidate>();
        var candidates = new PdfUnderstandingTableCandidate[result.Count];
        for (int index = 0; index < result.Count; index++) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            candidates[index] = result[index].Candidate;
        }
        return Array.AsReadOnly(candidates);
    }

    private static CandidateState CreateState(
        PdfReadPage page,
        PdfUnderstandingTableCandidate candidate,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        bool hasBounds = TryGetVisualBounds(page, candidate, consumeWork, cancellationCheck, out PdfVisualBounds bounds);
        return CreateState(candidate, hasBounds, bounds, consumeWork, cancellationCheck);
    }

    private static CandidateState CreateState(
        PdfLogicalPage page,
        PdfUnderstandingTableCandidate candidate,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        bool hasBounds = TryGetVisualBounds(page, candidate, consumeWork, cancellationCheck, out PdfVisualBounds bounds);
        return CreateState(candidate, hasBounds, bounds, consumeWork, cancellationCheck);
    }

    private static CandidateState CreateState(
        PdfUnderstandingTableCandidate candidate,
        bool hasBounds,
        PdfVisualBounds bounds,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        bool tagged = false;
        for (int evidenceIndex = 0; evidenceIndex < candidate.Evidence.Count; evidenceIndex++) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            if (!string.Equals(
                candidate.Evidence[evidenceIndex].Code,
                "table.tagged-structure",
                StringComparison.Ordinal)) continue;
            tagged = true;
            break;
        }

        var nativeSourceRuns = new HashSet<PdfTextSpan>();
        for (int runIndex = 0; runIndex < candidate.NativeSourceRuns.Count; runIndex++) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            nativeSourceRuns.Add(candidate.NativeSourceRuns[runIndex]);
        }

        Dictionary<string, int> rowSignatures = GetRowSignatures(
            candidate,
            consumeWork,
            cancellationCheck,
            out int rowSignatureCount,
            out int populatedCellCount);
        return new CandidateState(
            candidate,
            hasBounds,
            bounds,
            tagged,
            nativeSourceRuns,
            rowSignatures,
            rowSignatureCount,
            populatedCellCount);
    }

    private static CandidateRelationship GetRelationship(
        CandidateState left,
        CandidateState right,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        if (!left.HasBounds || !right.HasBounds) return CandidateRelationship.Distinct;
        PdfVisualBounds leftBounds = left.Bounds;
        PdfVisualBounds rightBounds = right.Bounds;
        double horizontalOverlap = Math.Max(0D, Math.Min(leftBounds.Right, rightBounds.Right) - Math.Max(leftBounds.Left, rightBounds.Left));
        double verticalOverlap = Math.Max(0D, Math.Min(leftBounds.Bottom, rightBounds.Bottom) - Math.Max(leftBounds.Top, rightBounds.Top));
        double narrowerWidth = Math.Min(leftBounds.Width, rightBounds.Width);
        double shorterHeight = Math.Min(leftBounds.Height, rightBounds.Height);
        if (narrowerWidth <= 0D || shorterHeight <= 0D) return CandidateRelationship.Distinct;

        double horizontalRatio = horizontalOverlap / narrowerWidth;
        double verticalRatio = verticalOverlap / shorterHeight;
        if (horizontalRatio < 0.5D || verticalRatio < 0.3D) return CandidateRelationship.Distinct;

        if (left.Tagged != right.Tagged) {
            if (right.Tagged && ContainsAllNativeSourceRuns(right, left, consumeWork, cancellationCheck)) {
                return CandidateRelationship.AdditionRicher;
            }
            if (left.Tagged && ContainsAllNativeSourceRuns(left, right, consumeWork, cancellationCheck)) {
                return CandidateRelationship.ExistingRicher;
            }
            if (right.Tagged &&
                ContainsAllNativeSourceRuns(left, right, consumeWork, cancellationCheck) &&
                HaveEquivalentRows(left, right, consumeWork, cancellationCheck)) {
                return CandidateRelationship.ExistingRicher;
            }
            if (left.Tagged &&
                ContainsAllNativeSourceRuns(right, left, consumeWork, cancellationCheck) &&
                HaveEquivalentRows(left, right, consumeWork, cancellationCheck)) {
                return CandidateRelationship.AdditionRicher;
            }
            // Partial tag ownership cannot prove that either candidate subsumes the other.
            // Keep both so untagged source content is never discarded by text similarity.
            return CandidateRelationship.Distinct;
        }

        if (HasStrictlyRicherContent(right, left) &&
            ContainsAll(right.RowSignatures, left.RowSignatures, consumeWork, cancellationCheck)) {
            return CandidateRelationship.AdditionRicher;
        }
        if (HasStrictlyRicherContent(left, right) &&
            ContainsAll(left.RowSignatures, right.RowSignatures, consumeWork, cancellationCheck)) {
            return CandidateRelationship.ExistingRicher;
        }
        if (left.RowSignatureCount == 0 || right.RowSignatureCount == 0) {
            return horizontalRatio >= 0.65D && verticalRatio >= 0.6D &&
                left.Candidate.Rows.Count == right.Candidate.Rows.Count
                ? CandidateRelationship.Duplicate
                : CandidateRelationship.Distinct;
        }
        int shared = CountShared(
            left.RowSignatures,
            right.RowSignatures,
            consumeWork,
            cancellationCheck);
        return shared >= 1 && shared * 2 > Math.Min(left.RowSignatureCount, right.RowSignatureCount)
            ? CandidateRelationship.Duplicate
            : CandidateRelationship.Distinct;
    }

    private static bool ContainsAllNativeSourceRuns(
        CandidateState container,
        CandidateState candidate,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        if (container.NativeSourceRuns.Count == 0 || candidate.NativeSourceRuns.Count == 0) return false;
        foreach (PdfTextSpan run in candidate.NativeSourceRuns) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            if (!container.NativeSourceRuns.Contains(run)) return false;
        }
        return true;
    }

    private static bool ContainsAll(
        Dictionary<string, int> candidate,
        Dictionary<string, int> existing,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        foreach (KeyValuePair<string, int> pair in existing) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            if (!candidate.TryGetValue(pair.Key, out int count) || count < pair.Value) return false;
        }
        return true;
    }

    private static bool HaveEquivalentRows(
        CandidateState left,
        CandidateState right,
        Action<long>? consumeWork,
        Action? cancellationCheck) =>
        left.RowSignatureCount == right.RowSignatureCount &&
        ContainsAll(left.RowSignatures, right.RowSignatures, consumeWork, cancellationCheck) &&
        ContainsAll(right.RowSignatures, left.RowSignatures, consumeWork, cancellationCheck);

    private static int CountShared(
        Dictionary<string, int> left,
        Dictionary<string, int> right,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        Dictionary<string, int> smaller = left.Count <= right.Count ? left : right;
        Dictionary<string, int> larger = ReferenceEquals(smaller, left) ? right : left;
        int shared = 0;
        foreach (KeyValuePair<string, int> pair in smaller) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            if (larger.TryGetValue(pair.Key, out int count)) shared += Math.Min(pair.Value, count);
        }
        return shared;
    }

    private static bool HasStrictlyRicherContent(CandidateState candidate, CandidateState existing) =>
        candidate.Candidate.Rows.Count > existing.Candidate.Rows.Count &&
        candidate.PopulatedCellCount > existing.PopulatedCellCount;

    private static bool TryGetVisualBounds(
        PdfLogicalPage page,
        PdfUnderstandingTableCandidate candidate,
        Action<long>? consumeWork,
        Action? cancellationCheck,
        out PdfVisualBounds bounds) {
        if (!TryGetCandidateBounds(
            candidate,
            consumeWork,
            cancellationCheck,
            out double left,
            out double right,
            out double yMinimum,
            out double yMaximum)) {
            bounds = default;
            return false;
        }
        bounds = candidate.CoordinateSpace == PdfTableCoordinateSpace.VisualTopLeft
            ? new PdfVisualBounds(left, yMinimum, right, yMaximum)
            : page.TransformBoundsToVisual(left, yMinimum, right, yMaximum);
        return bounds.Right > bounds.Left && bounds.Bottom > bounds.Top;
    }

    private static bool TryGetVisualBounds(
        PdfReadPage page,
        PdfUnderstandingTableCandidate candidate,
        Action<long>? consumeWork,
        Action? cancellationCheck,
        out PdfVisualBounds bounds) {
        if (!TryGetCandidateBounds(
            candidate,
            consumeWork,
            cancellationCheck,
            out double left,
            out double right,
            out double yMinimum,
            out double yMaximum)) {
            bounds = default;
            return false;
        }
        bounds = candidate.CoordinateSpace == PdfTableCoordinateSpace.VisualTopLeft
            ? new PdfVisualBounds(left, yMinimum, right, yMaximum)
            : page.TransformBoundsToVisual(left, yMinimum, right, yMaximum);
        return bounds.Right > bounds.Left && bounds.Bottom > bounds.Top;
    }

    private static bool TryGetCandidateBounds(
        PdfUnderstandingTableCandidate candidate,
        Action<long>? consumeWork,
        Action? cancellationCheck,
        out double left,
        out double right,
        out double yMinimum,
        out double yMaximum) {
        if (candidate.Columns.Count == 0) {
            left = right = yMinimum = yMaximum = 0D;
            return false;
        }
        if (candidate.CoordinateSpace == PdfTableCoordinateSpace.VisualTopLeft &&
            candidate.VisualBounds is PdfLogicalVisualBounds visual) {
            left = visual.Left;
            right = visual.Right;
            yMinimum = visual.Top;
            yMaximum = visual.Bottom;
            return true;
        }

        left = double.PositiveInfinity;
        right = double.NegativeInfinity;
        for (int columnIndex = 0; columnIndex < candidate.Columns.Count; columnIndex++) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            PdfUnderstandingTableColumn column = candidate.Columns[columnIndex];
            left = Math.Min(left, Math.Min(column.From, column.To));
            right = Math.Max(right, Math.Max(column.From, column.To));
        }
        yMinimum = Math.Min(candidate.YTop, candidate.YBottom);
        yMaximum = Math.Max(candidate.YTop, candidate.YBottom);
        return right > left && yMaximum > yMinimum;
    }

    private static Dictionary<string, int> GetRowSignatures(
        PdfUnderstandingTableCandidate candidate,
        Action<long>? consumeWork,
        Action? cancellationCheck,
        out int signatureCount,
        out int populatedCellCount) {
        var result = new Dictionary<string, int>(StringComparer.Ordinal);
        signatureCount = 0;
        populatedCellCount = 0;
        for (int rowIndex = 0; rowIndex < candidate.Rows.Count; rowIndex++) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            IReadOnlyList<string> row = candidate.Rows[rowIndex];
            var signature = new System.Text.StringBuilder();
            for (int columnIndex = 0; columnIndex < row.Count; columnIndex++) {
                cancellationCheck?.Invoke();
                consumeWork?.Invoke(1);
                if (columnIndex > 0) signature.Append('\u001F');
                string normalized = NormalizeCell(row[columnIndex], consumeWork, cancellationCheck);
                if (normalized.Length > 0) populatedCellCount++;
                signature.Append(normalized);
            }
            string value = signature.ToString();
            if (value.Trim('\u001F').Length == 0) continue;
            result[value] = result.TryGetValue(value, out int count) ? count + 1 : 1;
            signatureCount++;
        }
        return result;
    }

    private static string NormalizeCell(
        string value,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        string source = value.ToUpperInvariant();
        var result = new System.Text.StringBuilder(source.Length);
        bool pendingSpace = false;
        for (int index = 0; index < source.Length; index++) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
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

    private sealed class CandidateState {
        internal CandidateState(
            PdfUnderstandingTableCandidate candidate,
            bool hasBounds,
            PdfVisualBounds bounds,
            bool tagged,
            HashSet<PdfTextSpan> nativeSourceRuns,
            Dictionary<string, int> rowSignatures,
            int rowSignatureCount,
            int populatedCellCount) {
            Candidate = candidate;
            HasBounds = hasBounds;
            Bounds = bounds;
            Tagged = tagged;
            NativeSourceRuns = nativeSourceRuns;
            RowSignatures = rowSignatures;
            RowSignatureCount = rowSignatureCount;
            PopulatedCellCount = populatedCellCount;
        }

        internal PdfUnderstandingTableCandidate Candidate { get; }
        internal bool HasBounds { get; }
        internal PdfVisualBounds Bounds { get; }
        internal bool Tagged { get; }
        internal HashSet<PdfTextSpan> NativeSourceRuns { get; }
        internal Dictionary<string, int> RowSignatures { get; }
        internal int RowSignatureCount { get; }
        internal int PopulatedCellCount { get; }
    }

    private enum CandidateRelationship {
        Distinct,
        Duplicate,
        AdditionRicher,
        ExistingRicher
    }
}
