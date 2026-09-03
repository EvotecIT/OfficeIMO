namespace OfficeIMO.Pdf;

/// <summary>Coordinate system used by table bounds and columns.</summary>
public enum PdfTableCoordinateSpace {
    /// <summary>PDF user space with the Y axis increasing from bottom to top.</summary>
    PdfUserSpace,
    /// <summary>Top-left visual page space used by rendered OCR input.</summary>
    VisualTopLeft
}

/// <summary>Immutable column geometry recovered by a PDF table-detection stage.</summary>
public sealed class PdfUnderstandingTableColumn {
    /// <summary>Creates one table column in the owning candidate's coordinate space.</summary>
    public PdfUnderstandingTableColumn(double from, double to) {
        if (!IsFinite(from)) throw new ArgumentOutOfRangeException(nameof(from));
        if (!IsFinite(to)) throw new ArgumentOutOfRangeException(nameof(to));
        From = Math.Min(from, to);
        To = Math.Max(from, to);
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    /// <summary>Left X coordinate in the owning candidate's coordinate space.</summary>
    public double From { get; }

    /// <summary>Right X coordinate in the owning candidate's coordinate space.</summary>
    public double To { get; }
}

/// <summary>
/// A table candidate recovered before general page segmentation. The candidate owns its
/// source lines so later stages can treat the table as one structural object.
/// </summary>
public sealed class PdfUnderstandingTableCandidate {
    private readonly PdfLogicalVisualBounds? _visualBounds;

    /// <summary>Creates a table candidate for a custom table-detection stage.</summary>
    public PdfUnderstandingTableCandidate(
        string detectionKind,
        double yTop,
        double yBottom,
        IReadOnlyList<PdfUnderstandingTableColumn> columns,
        IReadOnlyList<IReadOnlyList<string>> rows,
        IReadOnlyList<PdfUnderstandingLine> sourceLines,
        double confidence = 0.5D,
        IEnumerable<PdfInferenceEvidence>? evidence = null)
        : this(
            detectionKind,
            yTop,
            yBottom,
            columns,
            rows,
            sourceLines,
            PdfLogicalContentSourceKind.Native,
            PdfTableCoordinateSpace.PdfUserSpace,
            null,
            confidence,
            evidence,
            null,
            null) {
    }

    private PdfUnderstandingTableCandidate(
        string detectionKind,
        double yTop,
        double yBottom,
        IReadOnlyList<PdfUnderstandingTableColumn> columns,
        IReadOnlyList<IReadOnlyList<string>> rows,
        IReadOnlyList<PdfUnderstandingLine> sourceLines,
        PdfLogicalContentSourceKind sourceKind,
        PdfTableCoordinateSpace coordinateSpace,
        PdfLogicalVisualBounds? visualBounds,
        double confidence,
        IEnumerable<PdfInferenceEvidence>? evidence,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        cancellationCheck?.Invoke();
        Guard.NotNull(detectionKind, nameof(detectionKind));
        Guard.NotNull(columns, nameof(columns));
        Guard.NotNull(rows, nameof(rows));
        Guard.NotNull(sourceLines, nameof(sourceLines));
        if (string.IsNullOrWhiteSpace(detectionKind)) throw new ArgumentException("Detection kind cannot be empty.", nameof(detectionKind));
        if (!IsFinite(yTop)) throw new ArgumentOutOfRangeException(nameof(yTop));
        if (!IsFinite(yBottom)) throw new ArgumentOutOfRangeException(nameof(yBottom));
        if (columns.Count < 2) throw new ArgumentException("A table candidate requires at least two columns.", nameof(columns));
        if (!IsFinite(confidence)) throw new ArgumentOutOfRangeException(nameof(confidence));

        DetectionKind = detectionKind;
        YTop = coordinateSpace == PdfTableCoordinateSpace.VisualTopLeft
            ? Math.Min(yTop, yBottom)
            : Math.Max(yTop, yBottom);
        YBottom = coordinateSpace == PdfTableCoordinateSpace.VisualTopLeft
            ? Math.Max(yTop, yBottom)
            : Math.Min(yTop, yBottom);
        Columns = SnapshotColumns(columns, consumeWork, cancellationCheck);
        Rows = SnapshotRows(rows, consumeWork, cancellationCheck);
        SourceLines = SnapshotSourceLines(sourceLines, consumeWork, cancellationCheck);
        SourceKind = sourceKind;
        CoordinateSpace = coordinateSpace;
        Confidence = PdfInference.Clamp(confidence);
        Evidence = PdfInference.Snapshot(evidence);
        _visualBounds = visualBounds;
    }

    /// <summary>Stable detector identifier that produced this candidate.</summary>
    public string DetectionKind { get; }

    /// <summary>Top Y coordinate in <see cref="CoordinateSpace"/>.</summary>
    public double YTop { get; private set; }

    /// <summary>Bottom Y coordinate in <see cref="CoordinateSpace"/>.</summary>
    public double YBottom { get; private set; }

    /// <summary>Detected columns in <see cref="CoordinateSpace"/>.</summary>
    public IReadOnlyList<PdfUnderstandingTableColumn> Columns { get; }

    /// <summary>Extracted row values aligned to <see cref="Columns"/>.</summary>
    public IReadOnlyList<IReadOnlyList<string>> Rows { get; }

    /// <summary>Exact native understanding-line fragments owned by this table.</summary>
    public IReadOnlyList<PdfUnderstandingLine> SourceLines { get; }

    /// <summary>Whether the candidate came from native PDF operations or accepted OCR geometry.</summary>
    public PdfLogicalContentSourceKind SourceKind { get; }

    /// <summary>Coordinate system used by <see cref="YTop"/>, <see cref="YBottom"/>, and <see cref="Columns"/>.</summary>
    public PdfTableCoordinateSpace CoordinateSpace { get; } = PdfTableCoordinateSpace.PdfUserSpace;

    /// <summary>Normalized table-detection confidence from 0 to 1.</summary>
    public double Confidence { get; }

    /// <summary>Evidence supporting the table candidate.</summary>
    public IReadOnlyList<PdfInferenceEvidence> Evidence { get; }

    internal PdfLogicalVisualBounds? VisualBounds => _visualBounds;

    internal static PdfUnderstandingTableCandidate FromStructured(
        StructuredTable table,
        IReadOnlyList<PdfUnderstandingLine> sourceLines,
        double confidence,
        IEnumerable<PdfInferenceEvidence> evidence,
        Action<long> consumeWork,
        Action cancellationCheck) {
        var columns = new PdfUnderstandingTableColumn[table.Columns.Count];
        for (int columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++) {
            cancellationCheck();
            consumeWork(1);
            StructuredTableColumn column = table.Columns[columnIndex];
            columns[columnIndex] = new PdfUnderstandingTableColumn(column.From, column.To);
        }
        var rows = new IReadOnlyList<string>[table.Rows.Count];
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            cancellationCheck();
            consumeWork(1);
            rows[rowIndex] = table.Rows[rowIndex];
        }
        return new PdfUnderstandingTableCandidate(
            table.Kind,
            table.YTop,
            table.YBottom,
            columns,
            rows,
            sourceLines,
            PdfLogicalContentSourceKind.Native,
            PdfTableCoordinateSpace.PdfUserSpace,
            null,
            confidence,
            evidence,
            consumeWork,
            cancellationCheck);
    }

    internal static PdfUnderstandingTableCandidate FromOcr(
        string detectionKind,
        double top,
        double bottom,
        PdfLogicalVisualBounds visualBounds,
        IReadOnlyList<(double From, double To)> visualColumnBounds,
        IReadOnlyList<IReadOnlyList<string>> rows,
        double confidence,
        IEnumerable<PdfInferenceEvidence> evidence) {
        var columns = visualColumnBounds
            .Select(static column => new PdfUnderstandingTableColumn(column.From, column.To))
            .ToArray();
        return new PdfUnderstandingTableCandidate(
            detectionKind,
            top,
            bottom,
            columns,
            rows,
            Array.Empty<PdfUnderstandingLine>(),
            PdfLogicalContentSourceKind.Ocr,
            PdfTableCoordinateSpace.VisualTopLeft,
            visualBounds,
            confidence,
            evidence,
            null,
            null);
    }

    internal StructuredTable ToStructuredTable(Action<long>? consumeWork = null, Action? cancellationCheck = null) {
        cancellationCheck?.Invoke();
        consumeWork?.Invoke(1);
        var table = new StructuredTable {
            Kind = DetectionKind,
            YTop = YTop,
            YBottom = YBottom
        };
        for (int columnIndex = 0; columnIndex < Columns.Count; columnIndex++) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            PdfUnderstandingTableColumn column = Columns[columnIndex];
            table.Columns.Add(new StructuredTableColumn { From = column.From, To = column.To });
        }
        for (int rowIndex = 0; rowIndex < Rows.Count; rowIndex++) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            IReadOnlyList<string> row = Rows[rowIndex];
            var cells = new string[row.Count];
            for (int columnIndex = 0; columnIndex < row.Count; columnIndex++) {
                cancellationCheck?.Invoke();
                consumeWork?.Invoke(1);
                cells[columnIndex] = row[columnIndex];
            }
            table.Rows.Add(cells);
        }
        return table;
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfUnderstandingTableColumn> SnapshotColumns(
        IReadOnlyList<PdfUnderstandingTableColumn> columns,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        var result = new PdfUnderstandingTableColumn[columns.Count];
        for (int index = 0; index < columns.Count; index++) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            result[index] = columns[index] ?? throw new ArgumentException("Columns cannot contain null values.", nameof(columns));
        }
        return Array.AsReadOnly(result);
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<IReadOnlyList<string>> SnapshotRows(
        IReadOnlyList<IReadOnlyList<string>> rows,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        var result = new IReadOnlyList<string>[rows.Count];
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            IReadOnlyList<string> row = rows[rowIndex] ?? throw new ArgumentException("Rows cannot contain null values.", nameof(rows));
            var cells = new string[row.Count];
            for (int columnIndex = 0; columnIndex < row.Count; columnIndex++) {
                cancellationCheck?.Invoke();
                consumeWork?.Invoke(1);
                cells[columnIndex] = row[columnIndex] ?? string.Empty;
            }
            result[rowIndex] = Array.AsReadOnly(cells);
        }
        return Array.AsReadOnly(result);
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfUnderstandingLine> SnapshotSourceLines(
        IReadOnlyList<PdfUnderstandingLine> sourceLines,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        var result = new PdfUnderstandingLine[sourceLines.Count];
        for (int index = 0; index < sourceLines.Count; index++) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            result[index] = sourceLines[index] ?? throw new ArgumentException("Source lines cannot contain null values.", nameof(sourceLines));
        }
        return Array.AsReadOnly(result);
    }
}
