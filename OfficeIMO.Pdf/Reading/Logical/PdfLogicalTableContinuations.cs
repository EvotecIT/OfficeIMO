namespace OfficeIMO.Pdf;

/// <summary>
/// Bounded continuation analysis shared by structured conversion adapters.
/// </summary>
internal static class PdfLogicalTableContinuations {
    private const int MaximumRepeatedHeaderRows = 4;

    internal static IReadOnlyList<PdfLogicalTableContinuationGroup> Group(
        PdfLogicalDocument document,
        int maxRows,
        bool mergePageContinuations,
        bool suppressRepeatedBodyHeaderRows,
        int maximumSegmentsPerTable,
        double geometryTolerancePoints) {
        Guard.NotNull(document, nameof(document));
#pragma warning disable CA1512 // ThrowIfLessThan is unavailable on netstandard2.0.
        if (maximumSegmentsPerTable < 1) throw new ArgumentOutOfRangeException(nameof(maximumSegmentsPerTable));
#pragma warning restore CA1512
        if (double.IsNaN(geometryTolerancePoints) || double.IsInfinity(geometryTolerancePoints) || geometryTolerancePoints < 0D) {
            throw new ArgumentOutOfRangeException(nameof(geometryTolerancePoints));
        }

        int extractionRowLimit = maxRows > 0
            ? maxRows > int.MaxValue - MaximumRepeatedHeaderRows
                ? int.MaxValue
                : maxRows + MaximumRepeatedHeaderRows
            : 0;
        IReadOnlyList<PdfLogicalTableExtraction> extractions =
            PdfLogicalTableAnalysis.ExtractTables(document, extractionRowLimit);
        if (extractions.Count == 0) return Array.Empty<PdfLogicalTableContinuationGroup>();

        var groups = new List<PdfLogicalTableContinuationGroup>(extractions.Count);
        var segments = new List<PdfLogicalTableExtraction>();
        for (int index = 0; index < extractions.Count; index++) {
            PdfLogicalTableExtraction current = extractions[index];
            if (segments.Count > 0 && (!mergePageContinuations ||
                segments.Count >= maximumSegmentsPerTable ||
                !CanContinue(document, segments[segments.Count - 1], current, geometryTolerancePoints))) {
                groups.Add(CreateGroup(segments, maxRows, suppressRepeatedBodyHeaderRows));
                segments.Clear();
            }

            segments.Add(current);
        }

        if (segments.Count > 0) groups.Add(CreateGroup(segments, maxRows, suppressRepeatedBodyHeaderRows));
        return groups.AsReadOnly();
    }

    private static bool CanContinue(
        PdfLogicalDocument document,
        PdfLogicalTableExtraction previous,
        PdfLogicalTableExtraction current,
        double tolerance) {
        if (current.PageIndex != previous.PageIndex + 1 || current.PageNumber != previous.PageNumber + 1) return false;
        PdfLogicalPage previousPage = document.Pages[previous.PageIndex];
        PdfLogicalPage currentPage = document.Pages[current.PageIndex];
        if (previous.TableIndex != previousPage.Tables.Count - 1 || current.TableIndex != 0) return false;
        if (previous.Data.Columns.Count < 2 || previous.Data.Columns.Count != current.Data.Columns.Count) return false;
        if (!string.Equals(previous.DetectionKind, current.DetectionKind, StringComparison.Ordinal)) return false;
        if (!IsAtBottomEdge(previous.Table, previousPage) || !IsAtTopEdge(current.Table, currentPage)) return false;
        if (!HasCompatibleColumns(previous.Table, current.Table, tolerance)) return false;

        bool previousHasHeader = previous.Data.Structure.HasHeaderRow;
        bool currentHasHeader = current.Data.Structure.HasHeaderRow;
        if (!currentHasHeader) return true;
        return previousHasHeader && HeadersEqual(previous.Data.Columns, current.Data.Columns);
    }

    private static bool IsAtBottomEdge(PdfLogicalTable table, PdfLogicalPage page) =>
        table.YBottom <= Math.Max(18D, page.Height * 0.28D);

    private static bool IsAtTopEdge(PdfLogicalTable table, PdfLogicalPage page) =>
        table.YTop >= page.Height * 0.72D;

    internal static bool HasCompatibleColumns(
        PdfLogicalTable previousTable,
        PdfLogicalTable currentTable,
        double tolerance) {
        IReadOnlyList<PdfLogicalTableColumn> previous = previousTable.Columns;
        IReadOnlyList<PdfLogicalTableColumn> current = currentTable.Columns;
        if (previous.Count == 0 || previous.Count != current.Count) return false;
        bool positionedRecovery = string.Equals(previousTable.DetectionKind, "positioned-cells-bounded", StringComparison.Ordinal) &&
            string.Equals(currentTable.DetectionKind, "positioned-cells-bounded", StringComparison.Ordinal);
        for (int index = 0; index < previous.Count; index++) {
            if (Math.Abs(previous[index].From - current[index].From) > tolerance) return false;
            // The last right edge is based on the widest text run on each page rather than a stable split.
            // Positioned-cell recovery derives every right edge from page-local text width, so its
            // stable compatibility evidence is the ordered set of column starts.
            if (!positionedRecovery && index < previous.Count - 1 && Math.Abs(previous[index].To - current[index].To) > tolerance) return false;
        }

        return true;
    }

    private static bool HeadersEqual(IReadOnlyList<string> previous, IReadOnlyList<string> current) {
        if (previous.Count != current.Count) return false;
        for (int index = 0; index < previous.Count; index++) {
            if (!string.Equals(previous[index].Trim(), current[index].Trim(), StringComparison.OrdinalIgnoreCase)) return false;
        }

        return true;
    }

    private static PdfLogicalTableContinuationGroup CreateGroup(
        IReadOnlyList<PdfLogicalTableExtraction> sourceSegments,
        int maxRows,
        bool suppressRepeatedBodyHeaderRows) {
        PdfLogicalTableExtraction[] segments = sourceSegments.ToArray();
        int repeatedBodyHeaderRows = suppressRepeatedBodyHeaderRows
            ? DetectRepeatedBodyHeaderRows(segments)
            : 0;
        IReadOnlyList<string> columns = BuildColumns(segments[0].Data.Columns, segments[0].Data.Rows, repeatedBodyHeaderRows);
        var allRows = new List<IReadOnlyList<string>>();
        int totalRowCount = 0;
        int suppressedRows = 0;
        for (int segmentIndex = 0; segmentIndex < segments.Length; segmentIndex++) {
            IReadOnlyList<IReadOnlyList<string>> rows = segments[segmentIndex].Data.Rows;
            int start = repeatedBodyHeaderRows;
            int availableRowCount = segments[segmentIndex].Data.TotalRowCount;
            totalRowCount = checked(totalRowCount + Math.Max(0, availableRowCount - start));
            if (segmentIndex > 0) suppressedRows += Math.Min(start, availableRowCount);
            for (int rowIndex = start; rowIndex < rows.Count; rowIndex++) {
                if (maxRows <= 0 || allRows.Count < maxRows) allRows.Add(rows[rowIndex]);
            }
        }

        return new PdfLogicalTableContinuationGroup(
            segments,
            columns,
            allRows.AsReadOnly(),
            totalRowCount,
            allRows.Count < totalRowCount,
            suppressedRows,
            repeatedBodyHeaderRows);
    }

    private static int DetectRepeatedBodyHeaderRows(PdfLogicalTableExtraction[] segments) {
        if (segments.Length < 2 ||
            segments.Any(static segment => !segment.Data.Structure.HasHeaderRow) ||
            segments.Skip(1).Any(segment => !HeadersEqual(segments[0].Data.Columns, segment.Data.Columns))) {
            return 0;
        }
        int maximum = Math.Min(MaximumRepeatedHeaderRows, segments.Min(static segment => segment.Data.Rows.Count));
        int repeated = 0;
        for (int rowIndex = 0; rowIndex < maximum; rowIndex++) {
            IReadOnlyList<string> candidate = segments[0].Data.Rows[rowIndex];
            if (!LooksLikeHeaderRow(candidate)) break;
            bool matches = true;
            for (int segmentIndex = 1; segmentIndex < segments.Length; segmentIndex++) {
                if (!RowsEqual(candidate, segments[segmentIndex].Data.Rows[rowIndex])) {
                    matches = false;
                    break;
                }
            }

            if (!matches) break;
            repeated++;
        }

        return repeated;
    }

    private static bool LooksLikeHeaderRow(IReadOnlyList<string> row) {
        if (row.Count < 2) return false;
        var values = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (int index = 0; index < row.Count; index++) {
            string value = row[index].Trim();
            if (value.Length == 0 || !values.Add(value) || PdfLogicalTableAnalysis.LooksLikeNumericValue(value)) return false;
        }

        return true;
    }

    private static bool RowsEqual(IReadOnlyList<string> left, IReadOnlyList<string> right) {
        if (left.Count != right.Count) return false;
        for (int index = 0; index < left.Count; index++) {
            if (!string.Equals(left[index].Trim(), right[index].Trim(), StringComparison.OrdinalIgnoreCase)) return false;
        }

        return true;
    }

    private static IReadOnlyList<string> BuildColumns(
        IReadOnlyList<string> primaryHeaders,
        IReadOnlyList<IReadOnlyList<string>> primaryRows,
        int additionalHeaderRows) {
        if (additionalHeaderRows == 0) return primaryHeaders;
        var columns = new string[primaryHeaders.Count];
        for (int columnIndex = 0; columnIndex < columns.Length; columnIndex++) {
            var parts = new List<string>(additionalHeaderRows + 1) { primaryHeaders[columnIndex].Trim() };
            for (int rowIndex = 0; rowIndex < additionalHeaderRows; rowIndex++) {
                string part = columnIndex < primaryRows[rowIndex].Count ? primaryRows[rowIndex][columnIndex].Trim() : string.Empty;
                if (part.Length > 0 && !parts.Contains(part, StringComparer.OrdinalIgnoreCase)) parts.Add(part);
            }

            columns[columnIndex] = string.Join(" / ", parts);
        }

        return Array.AsReadOnly(columns);
    }
}

internal sealed class PdfLogicalTableContinuationGroup {
    internal PdfLogicalTableContinuationGroup(
        IReadOnlyList<PdfLogicalTableExtraction> segments,
        IReadOnlyList<string> columns,
        IReadOnlyList<IReadOnlyList<string>> rows,
        int totalRowCount,
        bool truncated,
        int suppressedRepeatedHeaderRows,
        int additionalHeaderRowCount) {
        Segments = segments;
        Columns = columns;
        Rows = rows;
        TotalRowCount = totalRowCount;
        Truncated = truncated;
        SuppressedRepeatedHeaderRows = suppressedRepeatedHeaderRows;
        AdditionalHeaderRowCount = additionalHeaderRowCount;
    }

    internal IReadOnlyList<PdfLogicalTableExtraction> Segments { get; }
    internal PdfLogicalTableExtraction Primary => Segments[0];
    internal IReadOnlyList<string> Columns { get; }
    internal IReadOnlyList<IReadOnlyList<string>> Rows { get; }
    internal int TotalRowCount { get; }
    internal bool Truncated { get; }
    internal int SuppressedRepeatedHeaderRows { get; }
    internal int AdditionalHeaderRowCount { get; }
}
