namespace OfficeIMO.Pdf;

/// <summary>The source cell a table cell diagnostic belongs to, such as a worksheet cell.</summary>
internal readonly struct PdfTableCellDiagnosticAddress {
    internal PdfTableCellDiagnosticAddress(string scope, int row, int column) {
        Scope = scope ?? string.Empty;
        Row = row;
        Column = column;
    }

    /// <summary>The grid the cell belongs to, such as a worksheet name.</summary>
    internal string Scope { get; }

    /// <summary>One-based source row.</summary>
    internal int Row { get; }

    /// <summary>One-based source column.</summary>
    internal int Column { get; }
}

/// <summary>One summarized finding: a contiguous rectangle of source cells that share a code.</summary>
internal sealed class PdfTableCellDiagnosticRange {
    internal PdfTableCellDiagnosticRange(string code, string scope, int firstRow, int firstColumn, int lastRow, int lastColumn,
        int cellCount, IReadOnlyList<(int Row, int Column)> samples, string firstMessage,
        PdfLayoutDiagnosticKind kind, PdfConversionWarningSeverity severity) {
        Code = code; Scope = scope; FirstRow = firstRow; FirstColumn = firstColumn; LastRow = lastRow; LastColumn = lastColumn;
        CellCount = cellCount; Samples = samples; FirstMessage = firstMessage; Kind = kind; Severity = severity;
    }

    internal string Code { get; }
    internal string Scope { get; }
    internal int FirstRow { get; }
    internal int FirstColumn { get; }
    internal int LastRow { get; }
    internal int LastColumn { get; }
    internal int CellCount { get; }
    internal IReadOnlyList<(int Row, int Column)> Samples { get; }
    internal string FirstMessage { get; }
    internal PdfLayoutDiagnosticKind Kind { get; }
    internal PdfConversionWarningSeverity Severity { get; }
}

/// <summary>
/// Collects clipped table cell findings by source cell and reports each code once per contiguous rectangle of
/// cells, instead of once per cell. The owning converter formats the reported ranges.
/// </summary>
internal abstract class PdfTableCellDiagnosticSummary {
    internal const int SampleCount = 3;

    private readonly Dictionary<(string Code, string Scope), Group> _groups = new();
    private readonly List<(string Code, string Scope)> _order = new();

    /// <summary>Codes whose per-cell occurrences this summary replaces.</summary>
    internal static bool Summarizes(string code) =>
        string.Equals(code, "table-cell-content-clipped", StringComparison.Ordinal)
        || string.Equals(code, "table-cell-paint-clipped", StringComparison.Ordinal)
        || string.Equals(code, "table-cell-canvas-clip-paint-clipped", StringComparison.Ordinal);

    internal void Record(string code, PdfTableCellDiagnosticAddress address, string message,
        PdfLayoutDiagnosticKind kind, PdfConversionWarningSeverity severity) {
        var key = (code, address.Scope);
        if (!_groups.TryGetValue(key, out Group? group)) {
            group = new Group(kind, severity);
            _groups.Add(key, group);
            _order.Add(key);
        }
        // The first report for a cell is its finding; a later pass over the same cell adds nothing.
        if (!group.Cells.ContainsKey((address.Row, address.Column))) group.Cells.Add((address.Row, address.Column), message);
    }

    /// <summary>Returns the collected findings as ranges and starts over, so a later write reports its own findings.</summary>
    internal IReadOnlyList<PdfTableCellDiagnosticRange> Complete() {
        var ranges = new List<PdfTableCellDiagnosticRange>();
        foreach (var key in _order) {
            Group group = _groups[key];
            // Ranges are reported in reading order of their first cell.
            foreach (var rectangle in CreateRectangles(group.Cells.Keys).OrderBy(item => item.FirstRow).ThenBy(item => item.FirstColumn)) {
                // Rectangles are filled, so the first cells in reading order are the samples.
                var samples = new List<(int Row, int Column)>(SampleCount);
                for (int row = rectangle.FirstRow; row <= rectangle.LastRow && samples.Count < SampleCount; row++)
                    for (int column = rectangle.FirstColumn; column <= rectangle.LastColumn && samples.Count < SampleCount; column++)
                        samples.Add((row, column));
                ranges.Add(new PdfTableCellDiagnosticRange(key.Code, key.Scope, rectangle.FirstRow, rectangle.FirstColumn,
                    rectangle.LastRow, rectangle.LastColumn, rectangle.Count, samples, group.Cells[samples[0]], group.Kind, group.Severity));
            }
        }
        _groups.Clear();
        _order.Clear();
        return ranges;
    }

    /// <summary>Text that names a range for the report's location.</summary>
    internal abstract string FormatSource(PdfTableCellDiagnosticRange range);

    /// <summary>Text that tells the reader what happened in the range.</summary>
    internal abstract string FormatMessage(PdfTableCellDiagnosticRange range);

    /// <summary>
    /// Splits cells into rectangles: runs of consecutive rows within a column, then neighbouring columns with
    /// the same run joined into one rectangle. Every cell belongs to exactly one rectangle.
    /// </summary>
    internal static IReadOnlyList<(int FirstRow, int FirstColumn, int LastRow, int LastColumn, int Count)> CreateRectangles(
        IEnumerable<(int Row, int Column)> cells) {
        var runs = new List<(int Column, int FirstRow, int LastRow)>();
        foreach (var column in cells.GroupBy(cell => cell.Column).OrderBy(group => group.Key)) {
            int start = -1, previous = -1;
            foreach (int row in column.Select(cell => cell.Row).OrderBy(row => row)) {
                if (start >= 0 && row == previous + 1) { previous = row; continue; }
                if (start >= 0) runs.Add((column.Key, start, previous));
                start = previous = row;
            }
            if (start >= 0) runs.Add((column.Key, start, previous));
        }
        var rectangles = new List<(int FirstRow, int FirstColumn, int LastRow, int LastColumn, int Count)>();
        // Open rectangles keyed by row span; a run extends one only from the immediately preceding column.
        var open = new Dictionary<(int FirstRow, int LastRow), int>();
        int currentColumn = int.MinValue;
        var next = new Dictionary<(int FirstRow, int LastRow), int>();
        foreach (var run in runs) {
            if (run.Column != currentColumn) {
                bool adjacent = run.Column == currentColumn + 1;
                open = adjacent ? next : new Dictionary<(int, int), int>();
                next = new Dictionary<(int, int), int>();
                currentColumn = run.Column;
            }
            var span = (run.FirstRow, run.LastRow);
            int rows = run.LastRow - run.FirstRow + 1;
            if (open.TryGetValue(span, out int index)) {
                var rectangle = rectangles[index];
                rectangles[index] = (rectangle.FirstRow, rectangle.FirstColumn, rectangle.LastRow, run.Column, rectangle.Count + rows);
                next[span] = index;
            } else {
                rectangles.Add((run.FirstRow, run.Column, run.LastRow, run.Column, rows));
                next[span] = rectangles.Count - 1;
            }
        }
        return rectangles;
    }

    private sealed class Group {
        internal Group(PdfLayoutDiagnosticKind kind, PdfConversionWarningSeverity severity) {
            Kind = kind; Severity = severity;
        }

        internal PdfLayoutDiagnosticKind Kind { get; }
        internal PdfConversionWarningSeverity Severity { get; }
        // Each cell's own message, so a range can quote the finding of its first cell.
        internal Dictionary<(int Row, int Column), string> Cells { get; } = new();
    }
}
