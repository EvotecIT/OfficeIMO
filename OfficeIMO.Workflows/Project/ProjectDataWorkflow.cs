using System.Globalization;
using OfficeIMO.CSV;
using OfficeIMO.Excel;
using OfficeIMO.Project;

namespace OfficeIMO.Workflows;

/// <summary>CSV and Excel transport for explicit Project field maps. Parsing and package handling remain in their owning libraries.</summary>
public static class ProjectDataWorkflow {
    /// <summary>Creates a CSV document containing the exact table text. Save/load settings, quoting and delimiter handling belong to OfficeIMO.CSV.</summary>
    public static CsvDocument CreateCsv(ProjectDataTable table, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(table);
        cancellationToken.ThrowIfCancellationRequested();
        var csv = new CsvDocument().WithHeader(table.Headers.ToArray());
        foreach (var row in table.Rows) { cancellationToken.ThrowIfCancellationRequested(); csv.AddRow(row.Cast<object?>().ToArray()); }
        return csv;
    }

    /// <summary>Copies a parsed CSV table without reflection or culture-dependent value conversion.</summary>
    public static ProjectDataTable ReadCsv(CsvDocument csv, int maxRows = 100000, int maxCells = 2000000, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(csv);
        cancellationToken.ThrowIfCancellationRequested();
        IEnumerable<IReadOnlyList<string?>> Rows() {
            foreach (var row in csv.AsEnumerable()) {
                cancellationToken.ThrowIfCancellationRequested();
                var values = new string?[row.FieldCount];
                for (int i = 0; i < values.Length; i++) values[i] = Invariant(row[i]);
                yield return values;
            }
        }
        return new ProjectDataTable(csv.Header, Rows(), maxRows, maxCells);
    }

    /// <summary>Creates a workbook with one worksheet per mapped table and a separate sheet describing projection losses. Values remain literal text to preserve precision and prevent formula execution.</summary>
    public static ExcelDocument CreateExcel(ProjectDataExportResult export, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(export);
        cancellationToken.ThrowIfCancellationRequested();
        var document = ExcelDocument.Create();
        try {
            foreach (var mapped in export.Tables) {
                ValidateExcelGrid(mapped.Table.Rows.Count, mapped.Table.Headers.Count, mapped.Kind.ToString());
                var sheet = document.AddWorksheet(mapped.Kind.ToString());
                for (int c = 0; c < mapped.Table.Headers.Count; c++) sheet.CellValue(1, c + 1, mapped.Table.Headers[c]);
                for (int r = 0; r < mapped.Table.Rows.Count; r++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    for (int c = 0; c < mapped.Table.Headers.Count; c++) {
                        string? value = mapped.Table.Rows[r][c]; if (value != null) sheet.CellValue(r + 2, c + 1, value);
                    }
                }
                sheet.Freeze(1); sheet.AutoFitColumns(ct: cancellationToken);
            }
            ValidateExcelGrid(export.Notices.Count, 1, "Transfer notes");
            var notes = document.AddWorksheet("Transfer notes"); notes.CellValue(1, 1, "Projection boundaries");
            for (int i = 0; i < export.Notices.Count; i++) notes.CellValue(i + 2, 1, export.Notices[i]);
            return document;
        } catch { document.Dispose(); throw; }
    }

    internal static void ValidateExcelGrid(int dataRowCount, int columnCount, string tableName) {
        if (dataRowCount < 0 || dataRowCount > A1.MaxRows - 1)
            throw new InvalidOperationException($"Project table '{tableName}' exceeds Excel's {A1.MaxRows:N0}-row worksheet limit including its header.");
        if (columnCount < 1 || columnCount > A1.MaxColumns)
            throw new InvalidOperationException($"Project table '{tableName}' exceeds Excel's {A1.MaxColumns:N0}-column worksheet limit.");
    }

    /// <summary>Reads an explicitly bounded worksheet rectangle with headers in its first row. Formula and error cells are rejected; numeric/date cells use invariant underlying values.</summary>
    public static ProjectDataTable ReadExcel(ExcelSheet sheet, int dataRowCount, int columnCount, int headerRow = 1, int firstColumn = 1,
        int maxRows = 100000, int maxCells = 2000000, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(sheet);
        cancellationToken.ThrowIfCancellationRequested();
        if (dataRowCount < 0 || dataRowCount > maxRows || columnCount < 1 || columnCount > 64 || headerRow < 1 || firstColumn < 1
            || (long)headerRow + dataRowCount > 1048576 || (long)firstColumn + columnCount - 1 > 16384
            || (long)dataRowCount * columnCount > maxCells) throw new ArgumentOutOfRangeException(nameof(dataRowCount));
        string? Cell(int row, int column) {
            if (!sheet.TryGetCellValueSnapshot(row, column, out var value) || value == null) return null;
            return value.Kind switch {
                ExcelCellValueKind.Formula or ExcelCellValueKind.Error => throw new InvalidDataException("Formula or error cell at row " + row + ", column " + column + " cannot be imported as project data."),
                ExcelCellValueKind.DateTime => Invariant(value.DateTimeValue),
                ExcelCellValueKind.Boolean => value.RawValue == "1" ? "true" : "false",
                ExcelCellValueKind.Number => OfficeInvariantDecimal.TryParseExact(value.RawValue, true, out decimal numeric)
                    ? numeric.ToString("0.############################", CultureInfo.InvariantCulture)
                    : throw new InvalidDataException("Numeric cell at row " + row + ", column " + column + " cannot be represented exactly as a project decimal."),
                ExcelCellValueKind.Text => value.Text,
                _ => throw new InvalidDataException("Unsupported cell kind at row " + row + ", column " + column + ".")
            };
        }
        var headers = Enumerable.Range(firstColumn, columnCount).Select(c => Cell(headerRow, c) ?? "").ToArray();
        IEnumerable<IReadOnlyList<string?>> Rows() {
            for (int r = 0; r < dataRowCount; r++) {
                cancellationToken.ThrowIfCancellationRequested();
                yield return Enumerable.Range(firstColumn, columnCount).Select(c => Cell(headerRow + r + 1, c)).ToArray();
            }
        }
        return new ProjectDataTable(headers, Rows(), maxRows, maxCells);
    }

    private static string? Invariant(object? value) => value switch {
        null or DBNull => null,
        string text => text,
        DateTime date when date.Kind == DateTimeKind.Unspecified => date.ToString("yyyy-MM-ddTHH:mm:ss.fffffff", CultureInfo.InvariantCulture),
        DateTime => throw new InvalidDataException("Project table dates must explicitly use DateTimeKind.Unspecified; timezone-bearing values are not converted implicitly."),
        bool flag => flag ? "true" : "false",
        IFormattable formatted => formatted.ToString(null, CultureInfo.InvariantCulture),
        _ => throw new InvalidDataException("The input contains a value that has no portable project-table representation.")
    };
}
