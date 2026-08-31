namespace OfficeIMO.Pdf;

/// <summary>Controls bounded recovery of tables split across adjacent PDF pages.</summary>
public sealed class PdfLogicalTableContinuationOptions {
    /// <summary>Maximum merged body rows. Values less than or equal to zero retain all rows. Default: 0.</summary>
    public int MaxRows { get; init; }

    /// <summary>Whether adjacent page-edge table segments may be merged. Default: true.</summary>
    public bool MergePageContinuations { get; init; } = true;

    /// <summary>Whether repeated header-like body prefixes should be removed from continuation segments. Default: true.</summary>
    public bool SuppressRepeatedBodyHeaderRows { get; init; } = true;

    /// <summary>Maximum adjacent segments in one recovered table. Default: 64.</summary>
    public int MaximumSegmentsPerTable { get; init; } = 64;

    /// <summary>Maximum per-column geometry difference in visual PDF points. Default: 4.</summary>
    public double GeometryTolerancePoints { get; init; } = 4D;

    /// <summary>Minimum normalized continuation confidence required for a merge. Default: 0.75.</summary>
    public double MinimumConfidence { get; init; } = 0.75D;

    internal static PdfLogicalTableContinuationOptions Resolve(PdfLogicalTableContinuationOptions? options) {
        PdfLogicalTableContinuationOptions effective = options ?? new PdfLogicalTableContinuationOptions();
        if (effective.MaximumSegmentsPerTable < 1) {
            throw new ArgumentOutOfRangeException(nameof(options), effective.MaximumSegmentsPerTable, "Maximum table segments must be positive.");
        }
        if (double.IsNaN(effective.GeometryTolerancePoints) ||
            double.IsInfinity(effective.GeometryTolerancePoints) ||
            effective.GeometryTolerancePoints < 0D) {
            throw new ArgumentOutOfRangeException(nameof(options), effective.GeometryTolerancePoints, "Table geometry tolerance must be finite and nonnegative.");
        }
        if (double.IsNaN(effective.MinimumConfidence) ||
            double.IsInfinity(effective.MinimumConfidence) ||
            effective.MinimumConfidence < 0D ||
            effective.MinimumConfidence > 1D) {
            throw new ArgumentOutOfRangeException(nameof(options), effective.MinimumConfidence, "Table continuation confidence must be between zero and one.");
        }
        return effective;
    }
}

/// <summary>Evidence supporting one or more recovered cross-page table boundaries.</summary>
[Flags]
public enum PdfLogicalTableContinuationEvidence {
    /// <summary>No cross-page continuation was inferred.</summary>
    None = 0,
    /// <summary>The segments came from adjacent source pages.</summary>
    AdjacentPages = 1,
    /// <summary>The segments were the last and first table on their pages.</summary>
    BoundaryTables = 2,
    /// <summary>The segments were positioned near the bottom and top page edges.</summary>
    PageEdges = 4,
    /// <summary>The segments declared the same number of columns.</summary>
    MatchingColumnCount = 8,
    /// <summary>The segments came from the same table detection strategy.</summary>
    MatchingDetectionKind = 16,
    /// <summary>The ordered column geometry was compatible.</summary>
    CompatibleGeometry = 32,
    /// <summary>The continuation header contract was compatible.</summary>
    CompatibleHeaders = 64,
    /// <summary>Repeated continuation headers were detected.</summary>
    RepeatedHeaders = 128
}

/// <summary>
/// Bounded continuation analysis shared by structured conversion adapters.
/// </summary>
public static class PdfLogicalTableContinuations {
    private const int MaximumRepeatedHeaderRows = 4;

    /// <summary>
    /// Groups compatible table segments on adjacent pages and returns normalized merged data.
    /// </summary>
    /// <param name="document">Logical PDF document to analyze.</param>
    /// <param name="maxRows">Maximum merged body rows. Values less than or equal to zero retain all rows.</param>
    /// <param name="mergePageContinuations">Whether adjacent page-edge segments may be merged.</param>
    /// <param name="suppressRepeatedBodyHeaderRows">Whether repeated header-like body prefixes should be suppressed.</param>
    /// <param name="maximumSegmentsPerTable">Maximum adjacent segments in one group.</param>
    /// <param name="geometryTolerancePoints">Maximum per-column geometry difference in PDF points.</param>
    public static IReadOnlyList<PdfLogicalTableContinuationGroup> Group(
        PdfLogicalDocument document,
        int maxRows,
        bool mergePageContinuations,
        bool suppressRepeatedBodyHeaderRows,
        int maximumSegmentsPerTable,
        double geometryTolerancePoints) {
        return Group(document, new PdfLogicalTableContinuationOptions {
            MaxRows = maxRows,
            MergePageContinuations = mergePageContinuations,
            SuppressRepeatedBodyHeaderRows = suppressRepeatedBodyHeaderRows,
            MaximumSegmentsPerTable = maximumSegmentsPerTable,
            GeometryTolerancePoints = geometryTolerancePoints,
            MinimumConfidence = 0D
        });
    }

    /// <summary>Groups compatible table segments using an explicit, typed continuation policy.</summary>
    public static IReadOnlyList<PdfLogicalTableContinuationGroup> Group(
        PdfLogicalDocument document,
        PdfLogicalTableContinuationOptions? options = null) {
        Guard.NotNull(document, nameof(document));
        PdfLogicalTableContinuationOptions effective = PdfLogicalTableContinuationOptions.Resolve(options);

        int extractionRowLimit = effective.MaxRows > 0
            ? effective.MaxRows > int.MaxValue - MaximumRepeatedHeaderRows
                ? int.MaxValue
                : effective.MaxRows + MaximumRepeatedHeaderRows
            : 0;
        IReadOnlyList<PdfLogicalTableExtraction> extractions =
            PdfLogicalTableAnalysis.ExtractTables(document, extractionRowLimit);
        if (extractions.Count == 0) return Array.Empty<PdfLogicalTableContinuationGroup>();

        var groups = new List<PdfLogicalTableContinuationGroup>(extractions.Count);
        var segments = new List<PdfLogicalTableExtraction>();
        var boundaries = new List<ContinuationBoundary>();
        for (int index = 0; index < extractions.Count; index++) {
            PdfLogicalTableExtraction current = extractions[index];
            if (segments.Count > 0) {
                ContinuationBoundary boundary = default;
                bool canContinue = effective.MergePageContinuations &&
                    segments.Count < effective.MaximumSegmentsPerTable &&
                    CanContinue(document, segments[segments.Count - 1], current, effective.GeometryTolerancePoints, out boundary) &&
                    boundary.Confidence >= effective.MinimumConfidence;
                if (canContinue) {
                    boundaries.Add(boundary);
                } else {
                    groups.Add(CreateGroup(segments, boundaries, effective.MaxRows, effective.SuppressRepeatedBodyHeaderRows));
                    segments.Clear();
                    boundaries.Clear();
                }
            }

            segments.Add(current);
        }

        if (segments.Count > 0) groups.Add(CreateGroup(segments, boundaries, effective.MaxRows, effective.SuppressRepeatedBodyHeaderRows));
        return groups.AsReadOnly();
    }

    private static bool CanContinue(
        PdfLogicalDocument document,
        PdfLogicalTableExtraction previous,
        PdfLogicalTableExtraction current,
        double tolerance,
        out ContinuationBoundary boundary) {
        boundary = default;
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
        if (currentHasHeader && (!previousHasHeader || !HeadersEqual(previous.Data.Columns, current.Data.Columns))) return false;

        PdfLogicalTableContinuationEvidence evidence =
            PdfLogicalTableContinuationEvidence.AdjacentPages |
            PdfLogicalTableContinuationEvidence.BoundaryTables |
            PdfLogicalTableContinuationEvidence.PageEdges |
            PdfLogicalTableContinuationEvidence.MatchingColumnCount |
            PdfLogicalTableContinuationEvidence.MatchingDetectionKind |
            PdfLogicalTableContinuationEvidence.CompatibleGeometry |
            PdfLogicalTableContinuationEvidence.CompatibleHeaders;
        double confidence = 0.75D;
        if (HasCompatibleColumns(previous.Table, current.Table, tolerance * 0.5D)) confidence += 0.1D;
        if (currentHasHeader) {
            evidence |= PdfLogicalTableContinuationEvidence.RepeatedHeaders;
            confidence += 0.1D;
        } else {
            confidence += 0.05D;
        }
        boundary = new ContinuationBoundary(Math.Min(1D, confidence), evidence);
        return true;
    }

    private static bool IsAtBottomEdge(PdfLogicalTable table, PdfLogicalPage page) {
        if (!TryGetVisualBounds(table, page, out PdfVisualBounds bounds)) return false;
        (_, double visualHeight) = page.GetVisualPageSize();
        return bounds.Bottom >= visualHeight * 0.72D;
    }

    private static bool IsAtTopEdge(PdfLogicalTable table, PdfLogicalPage page) {
        if (!TryGetVisualBounds(table, page, out PdfVisualBounds bounds)) return false;
        (_, double visualHeight) = page.GetVisualPageSize();
        return bounds.Top <= Math.Max(18D, visualHeight * 0.28D);
    }

    private static bool TryGetVisualBounds(PdfLogicalTable table, PdfLogicalPage page, out PdfVisualBounds bounds) {
        if (table.Columns.Count == 0) { bounds = default; return false; }
        double left = table.Columns.Min(static column => Math.Min(column.From, column.To));
        double right = table.Columns.Max(static column => Math.Max(column.From, column.To));
        double bottom = Math.Min(table.YBottom, table.YTop);
        double top = Math.Max(table.YBottom, table.YTop);
        if (right <= left || top <= bottom) { bounds = default; return false; }
        bounds = page.TransformBoundsToVisual(left, bottom, right, top);
        return bounds.Right > bounds.Left && bounds.Bottom > bounds.Top;
    }

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
        IReadOnlyList<ContinuationBoundary> sourceBoundaries,
        int maxRows,
        bool suppressRepeatedBodyHeaderRows) {
        PdfLogicalTableExtraction[] segments = sourceSegments.ToArray();
        int repeatedBodyHeaderRows = suppressRepeatedBodyHeaderRows
            ? DetectRepeatedBodyHeaderRows(segments)
            : 0;
        double confidence = sourceBoundaries.Count == 0 ? 1D : sourceBoundaries.Min(static boundary => boundary.Confidence);
        PdfLogicalTableContinuationEvidence evidence = PdfLogicalTableContinuationEvidence.None;
        for (int boundaryIndex = 0; boundaryIndex < sourceBoundaries.Count; boundaryIndex++) {
            evidence |= sourceBoundaries[boundaryIndex].Evidence;
        }
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
            repeatedBodyHeaderRows,
            confidence,
            evidence);
    }

    private readonly struct ContinuationBoundary {
        internal ContinuationBoundary(double confidence, PdfLogicalTableContinuationEvidence evidence) {
            Confidence = confidence;
            Evidence = evidence;
        }

        internal double Confidence { get; }
        internal PdfLogicalTableContinuationEvidence Evidence { get; }
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

/// <summary>One logical table reconstructed from one or more adjacent page-level segments.</summary>
public sealed class PdfLogicalTableContinuationGroup {
    internal PdfLogicalTableContinuationGroup(
        IReadOnlyList<PdfLogicalTableExtraction> segments,
        IReadOnlyList<string> columns,
        IReadOnlyList<IReadOnlyList<string>> rows,
        int totalRowCount,
        bool truncated,
        int suppressedRepeatedHeaderRows,
        int additionalHeaderRowCount,
        double confidence,
        PdfLogicalTableContinuationEvidence evidence) {
        Segments = segments;
        Columns = columns;
        Rows = rows;
        TotalRowCount = totalRowCount;
        Truncated = truncated;
        SuppressedRepeatedHeaderRows = suppressedRepeatedHeaderRows;
        AdditionalHeaderRowCount = additionalHeaderRowCount;
        Confidence = confidence;
        Evidence = evidence;
        Data = CreateData(segments[0].Data, columns, rows, totalRowCount, truncated);
    }

    /// <summary>Page-level table segments contributing to this logical table.</summary>
    public IReadOnlyList<PdfLogicalTableExtraction> Segments { get; }
    /// <summary>Primary segment supplying source identity and diagnostics.</summary>
    public PdfLogicalTableExtraction Primary => Segments[0];
    /// <summary>Merged column names, including repeated multi-row header labels when requested.</summary>
    public IReadOnlyList<string> Columns { get; }
    /// <summary>Merged normalized body rows.</summary>
    public IReadOnlyList<IReadOnlyList<string>> Rows { get; }
    /// <summary>Normalized table data for downstream adapters.</summary>
    public PdfLogicalTableData Data { get; }
    /// <summary>Total body rows before the configured cap.</summary>
    public int TotalRowCount { get; }
    /// <summary>True when the configured row cap omitted body rows.</summary>
    public bool Truncated { get; }
    /// <summary>Repeated continuation header rows omitted from body data.</summary>
    public int SuppressedRepeatedHeaderRows { get; }
    /// <summary>Repeated header rows appended to the primary column labels.</summary>
    public int AdditionalHeaderRowCount { get; }
    /// <summary>Lowest normalized confidence across recovered page boundaries, or 1 for an unmerged table.</summary>
    public double Confidence { get; }
    /// <summary>Combined evidence supporting the recovered page boundaries.</summary>
    public PdfLogicalTableContinuationEvidence Evidence { get; }
    /// <summary>True when this logical table combines more than one page-level segment.</summary>
    public bool SpansPages => Segments.Count > 1;
    /// <summary>One-based source page number of the first segment.</summary>
    public int FirstPageNumber => Segments[0].PageNumber;
    /// <summary>One-based source page number of the last segment.</summary>
    public int LastPageNumber => Segments[Segments.Count - 1].PageNumber;

    private static PdfLogicalTableData CreateData(
        PdfLogicalTableData primary,
        IReadOnlyList<string> columns,
        IReadOnlyList<IReadOnlyList<string>> rows,
        int totalRowCount,
        bool truncated) {
        var numericColumns = new bool[columns.Count];
        for (int columnIndex = 0; columnIndex < columns.Count; columnIndex++) {
            bool hasValue = false;
            bool numeric = true;
            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
                string value = columnIndex < rows[rowIndex].Count ? rows[rowIndex][columnIndex] : string.Empty;
                if (string.IsNullOrWhiteSpace(value)) continue;
                hasValue = true;
                if (!PdfLogicalTableAnalysis.LooksLikeNumericValue(value)) { numeric = false; break; }
            }
            numericColumns[columnIndex] = hasValue && numeric;
        }
        var structure = new PdfLogicalTableStructure(
            columns.Count,
            columns,
            bodyStartRowIndex: 0,
            totalBodyRowCount: totalRowCount,
            hasHeaderRow: primary.Structure.HasHeaderRow,
            isKeyValueTable: primary.Structure.IsKeyValueTable);
        return new PdfLogicalTableData(structure, primary.Diagnostics, rows, numericColumns, truncated);
    }
}

public sealed partial class PdfLogicalDocument {
    /// <summary>Returns bounded cross-page table continuation groups in document order.</summary>
    public IReadOnlyList<PdfLogicalTableContinuationGroup> GetTableContinuationGroups(
        PdfLogicalTableContinuationOptions? options = null) =>
        PdfLogicalTableContinuations.Group(this, options);
}
