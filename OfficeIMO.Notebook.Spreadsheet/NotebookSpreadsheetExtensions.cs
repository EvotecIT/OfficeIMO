using System.Diagnostics;
using OfficeIMO.CSV;
using OfficeIMO.Excel;

namespace OfficeIMO.Notebook.Spreadsheet;

/// <summary>Exports a selected saved structured notebook output as tabular data.</summary>
public static class NotebookSpreadsheetExtensions {
    /// <summary>Saves the table as CSV to a caller-owned stream, escaping formula-like text by default.</summary>
    public static OfficeConversionReport SaveAsCsv(this NotebookOutput output, Stream destination,
        CsvSaveOptions? options = null, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        NotebookTable table = NotebookTable.Read(output, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        if (options?.ConfiguredFormulaInjectionPolicy == null) {
            options = options?.Clone() ?? new CsvSaveOptions();
            options.FormulaInjectionPolicy = CsvFormulaInjectionPolicy.Escape;
        }
        using (CsvRowWriter csv = CsvRowWriter.CreateStream(destination, options, leaveOpen: true)) {
            csv.WriteRows(table.Headers, ReadCsvRows(table, cancellationToken));
        }
        return new OfficeConversionReport("Jupyter Notebook output", "CSV", itemsWritten: table.RowCount, elapsed: timer.Elapsed);
    }

    /// <summary>Saves a one-sheet XLSX to a caller-owned stream through the Excel streaming writer.</summary>
    public static OfficeConversionReport SaveAsExcel(this NotebookOutput output, Stream destination,
        CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        NotebookTable table = NotebookTable.Read(output, cancellationToken);
        using (ExcelWorkbookWriter workbook = ExcelDocument.CreateWriter(destination, leaveOpen: true)) {
            workbook.WriteTable("Notebook data", table.Rows.EnumerateArray(), table.Headers, (writer, row) => {
                foreach (object? value in table.ReadRow(row, excel: true)) {
                    if (value == null) writer.WriteBlank();
                    else writer.Write(value);
                }
            }, cancellationToken);
        }
        return new OfficeConversionReport("Jupyter Notebook output", "XLSX", itemsWritten: table.RowCount, elapsed: timer.Elapsed);
    }

    private static IEnumerable<object?[]> ReadCsvRows(NotebookTable table, CancellationToken token) {
        foreach (var row in table.Rows.EnumerateArray()) {
            token.ThrowIfCancellationRequested();
            yield return table.ReadRow(row, excel: false);
        }
    }
}
