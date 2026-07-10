namespace OfficeIMO.OpenDocument;

internal static partial class OdfValidator {
    private static void ValidateSpreadsheetCell(XElement cell, string? valueType, List<OdfDiagnostic> diagnostics) {
        string? formula = (string?)cell.Attribute(OdfNamespaces.Table + "formula");
        if (formula != null && (string.IsNullOrWhiteSpace(formula) ||
            (!formula.Contains(":=") && !formula.StartsWith("=", StringComparison.Ordinal)))) {
            diagnostics.Add(new OdfDiagnostic("ODS102", OdfDiagnosticSeverity.Error,
                $"Spreadsheet formula '{formula}' does not contain a valid formula prefix.", "content.xml"));
        }

        string? lexical = null;
        bool valid = true;
        switch (valueType) {
            case "float":
            case "percentage":
            case "currency":
                lexical = (string?)cell.Attribute(OdfNamespaces.Office + "value");
                valid = lexical == null || double.TryParse(lexical, NumberStyles.Float, CultureInfo.InvariantCulture, out _);
                break;
            case "boolean":
                lexical = (string?)cell.Attribute(OdfNamespaces.Office + "boolean-value");
                valid = lexical == null || lexical == "true" || lexical == "false";
                break;
            case "date":
                lexical = (string?)cell.Attribute(OdfNamespaces.Office + "date-value");
                valid = lexical == null || DateTimeOffset.TryParse(lexical, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out _) ||
                    DateTime.TryParse(lexical, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out _);
                break;
            case "time":
                lexical = (string?)cell.Attribute(OdfNamespaces.Office + "time-value");
                if (lexical != null) {
                    try { _ = XmlConvert.ToTimeSpan(lexical); } catch (FormatException) { valid = false; }
                }
                break;
        }
        if (!valid) {
            diagnostics.Add(new OdfDiagnostic("ODS103", OdfDiagnosticSeverity.Error,
                $"Spreadsheet value '{lexical}' is not valid for value type '{valueType}'.", "content.xml"));
        }
    }

    private static void ValidateSpreadsheetMerges(XDocument content, List<OdfDiagnostic> diagnostics) {
        foreach (XElement table in content.Descendants(OdfNamespaces.Table + "table")) {
            List<ValidationRowRun> rows = BuildRowRuns(table);
            foreach (ValidationRowRun row in rows) {
                foreach (ValidationCellRun cell in row.Cells.Where(item => !item.Covered && (item.RowSpan > 1 || item.ColumnSpan > 1))) {
                    long mergeEndRow = SaturatingAdd(row.Start, cell.RowSpan);
                    long mergeEndColumn = SaturatingAdd(cell.Start, cell.ColumnSpan);
                    if (mergeEndRow == long.MaxValue || mergeEndColumn == long.MaxValue) {
                        AddMergeDiagnostic(diagnostics, "Merged cell span exceeds the supported coordinate range.");
                        continue;
                    }
                    if (cell.Repeat > 1) {
                        AddMergeDiagnostic(diagnostics, "Merged cell anchors cannot use number-columns-repeated.");
                        continue;
                    }
                    if (cell.ColumnSpan > 1 && !IsCovered(row.Cells, cell.Start + 1, mergeEndColumn)) {
                        AddMergeDiagnostic(diagnostics, "Merged cell column span is not followed by the required covered cells.");
                    }
                    if (cell.RowSpan > 1 && !RowsCoverMerge(rows, row.Start + 1, mergeEndRow, cell.Start, mergeEndColumn)) {
                        AddMergeDiagnostic(diagnostics, "Merged cell row span is not backed by the required covered cells.");
                    }
                }
            }
        }
    }

    private static List<ValidationRowRun> BuildRowRuns(XElement table) {
        var result = new List<ValidationRowRun>();
        long rowStart = 0;
        foreach (XElement row in table.Elements(OdfNamespaces.Table + "table-row")) {
            long rowRepeat = ReadPositive(row, OdfNamespaces.Table + "number-rows-repeated");
            var cells = new List<ValidationCellRun>();
            long columnStart = 0;
            foreach (XElement cell in OdsSheet.CellElements(row)) {
                long repeat = ReadPositive(cell, OdfNamespaces.Table + "number-columns-repeated");
                cells.Add(new ValidationCellRun(columnStart, repeat,
                    cell.Name == OdfNamespaces.Table + "covered-table-cell",
                    ReadPositive(cell, OdfNamespaces.Table + "number-rows-spanned"),
                    ReadPositive(cell, OdfNamespaces.Table + "number-columns-spanned")));
                columnStart = SaturatingAdd(columnStart, repeat);
            }
            result.Add(new ValidationRowRun(rowStart, rowRepeat, cells));
            rowStart = SaturatingAdd(rowStart, rowRepeat);
        }
        return result;
    }

    private static bool RowsCoverMerge(IReadOnlyList<ValidationRowRun> rows, long startRow, long endRow,
        long startColumn, long endColumn) {
        long cursor = startRow;
        foreach (ValidationRowRun row in rows.Where(item => item.Start < endRow && SaturatingAdd(item.Start, item.Repeat) > startRow)) {
            long overlapStart = Math.Max(cursor, row.Start);
            if (overlapStart > cursor || !IsCovered(row.Cells, startColumn, endColumn)) return false;
            cursor = Math.Min(endRow, SaturatingAdd(row.Start, row.Repeat));
            if (cursor >= endRow) return true;
        }
        return cursor >= endRow;
    }

    private static bool IsCovered(IReadOnlyList<ValidationCellRun> cells, long start, long end) {
        if (start >= end) return true;
        long cursor = start;
        foreach (ValidationCellRun cell in cells.Where(item => item.Start < end && SaturatingAdd(item.Start, item.Repeat) > start)) {
            if (cell.Start > cursor || !cell.Covered) return false;
            cursor = Math.Max(cursor, Math.Min(end, SaturatingAdd(cell.Start, cell.Repeat)));
            if (cursor >= end) return true;
        }
        return false;
    }

    private static long ReadPositive(XElement element, XName attribute) {
        string? lexical = (string?)element.Attribute(attribute);
        return lexical == null || !long.TryParse(lexical, NumberStyles.None, CultureInfo.InvariantCulture, out long value) || value < 1 ? 1 : value;
    }

    private static long SaturatingAdd(long left, long right) => right > long.MaxValue - left ? long.MaxValue : left + right;

    private static void AddMergeDiagnostic(List<OdfDiagnostic> diagnostics, string message) =>
        diagnostics.Add(new OdfDiagnostic("ODS104", OdfDiagnosticSeverity.Error, message, "content.xml"));

    private sealed class ValidationRowRun {
        internal ValidationRowRun(long start, long repeat, IReadOnlyList<ValidationCellRun> cells) {
            Start = start; Repeat = repeat; Cells = cells;
        }
        internal long Start { get; }
        internal long Repeat { get; }
        internal IReadOnlyList<ValidationCellRun> Cells { get; }
    }

    private sealed class ValidationCellRun {
        internal ValidationCellRun(long start, long repeat, bool covered, long rowSpan, long columnSpan) {
            Start = start; Repeat = repeat; Covered = covered; RowSpan = rowSpan; ColumnSpan = columnSpan;
        }
        internal long Start { get; }
        internal long Repeat { get; }
        internal bool Covered { get; }
        internal long RowSpan { get; }
        internal long ColumnSpan { get; }
    }
}
