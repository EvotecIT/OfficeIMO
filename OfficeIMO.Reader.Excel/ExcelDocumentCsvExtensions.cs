using System.Data;
using OfficeIMO.CSV;
using OfficeIMO.Excel;

namespace OfficeIMO.Reader.Excel;

/// <summary>Imports CSV and TSV content into Excel workbooks.</summary>
public static class ExcelDocumentCsvExtensions {
    /// <summary>Imports CSV or TSV text into a worksheet.</summary>
    public static ExcelDelimitedImportResult ImportDelimitedText(this ExcelDocument document, string text, ExcelDelimitedImportOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (text == null) throw new ArgumentNullException(nameof(text));
        options ??= new ExcelDelimitedImportOptions();
        int recordsToSkip = ValidateRecordsToSkip(options);
        char delimiter = options.Delimiter ?? DetectDelimiter(text, recordsToSkip);
        DataTable table = ExcelCsvDataTableBuilder.FromText(
            text,
            delimiter,
            options.HeadersInFirstRow,
            recordsToSkip,
            options.Culture,
            options.ConvertNumbersAndDates);
        return ImportTable(document, table, delimiter, options);
    }

    /// <summary>Imports a CSV or TSV file into a worksheet.</summary>
    public static ExcelDelimitedImportResult ImportDelimitedFile(this ExcelDocument document, string path, ExcelDelimitedImportOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        options ??= new ExcelDelimitedImportOptions();
        int recordsToSkip = ValidateRecordsToSkip(options);
        char delimiter = options.Delimiter ?? DetectDelimiterFromFile(path, recordsToSkip);
        DataTable table = ExcelCsvDataTableBuilder.FromFile(
            path,
            delimiter,
            options.HeadersInFirstRow,
            recordsToSkip,
            options.Culture,
            options.ConvertNumbersAndDates);
        return ImportTable(document, table, delimiter, options);
    }

    private static ExcelDelimitedImportResult ImportTable(ExcelDocument document, DataTable table, char delimiter, ExcelDelimitedImportOptions options) {
        table.TableName = string.IsNullOrWhiteSpace(options.SheetName) ? "Import" : options.SheetName!.Trim();
        var dataSet = new DataSet();
        dataSet.Tables.Add(table);
        ExcelDataSetImportResult imported = document.InsertDataSet(
            dataSet,
            createTables: false,
            tableStyle: options.TableStyle,
            includeHeaders: true,
            includeAutoFilter: true,
            autoFit: false)[0];

        string? actualTableName = null;
        if (options.CreateTable && !string.IsNullOrWhiteSpace(imported.Range)) {
            ExcelSheet sheet = document[imported.SheetName];
            string requestedTableName = string.IsNullOrWhiteSpace(options.TableName) ? imported.SheetName : options.TableName!.Trim();
            sheet.AddTable(imported.Range, hasHeader: true, requestedTableName, options.TableStyle, includeAutoFilter: true);
            actualTableName = document.GetTables()
                .FirstOrDefault(item => string.Equals(item.SheetName, imported.SheetName, StringComparison.OrdinalIgnoreCase) &&
                                        string.Equals(item.Range, imported.Range, StringComparison.OrdinalIgnoreCase))
                ?.Name;
        }

        return new ExcelDelimitedImportResult(
            imported.SheetName,
            actualTableName,
            imported.Range,
            imported.RowCount,
            imported.ColumnCount,
            delimiter,
            Array.Empty<string>());
    }

    private static int ValidateRecordsToSkip(ExcelDelimitedImportOptions options) {
        if (options.SkipInitialRecords < 0) {
            throw new ArgumentOutOfRangeException(nameof(options), "SkipInitialRecords cannot be negative.");
        }

        return options.SkipInitialRecords;
    }

    private static char DetectDelimiterFromFile(string path, int recordsToSkip) {
        using var reader = CsvFile.OpenTextReader(path);
        return DetectDelimiter(ReadFirstLogicalRecord(reader, recordsToSkip));
    }

    private static char DetectDelimiter(string text, int recordsToSkip) {
        using var reader = new StringReader(text);
        return DetectDelimiter(ReadFirstLogicalRecord(reader, recordsToSkip));
    }

    private static char DetectDelimiter(string record) {
        var candidates = new[] { ',', ';', '\t', '|' };
        return candidates
            .Select(candidate => new { Delimiter = candidate, Count = CountUnquoted(record, candidate) })
            .OrderByDescending(item => item.Count)
            .First()
            .Delimiter;
    }

    private static string ReadFirstLogicalRecord(TextReader reader, int recordsToSkip) {
        foreach (string record in ReadLogicalRecords(reader)) {
            if (record.Length == 0) continue;
            if (recordsToSkip > 0) {
                recordsToSkip--;
                continue;
            }

            return record;
        }

        return string.Empty;
    }

    private static IEnumerable<string> ReadLogicalRecords(TextReader reader) {
        string? line;
        while ((line = reader.ReadLine()) != null) {
            if (IsLogicalRecordComplete(line)) {
                yield return line;
                continue;
            }

            var record = new System.Text.StringBuilder(line);
            while ((line = reader.ReadLine()) != null) {
                record.Append('\n').Append(line);
                if (IsLogicalRecordComplete(record.ToString())) break;
            }

            yield return record.ToString();
        }
    }

    private static bool IsLogicalRecordComplete(string record) {
        bool quoted = false;
        for (int index = 0; index < record.Length; index++) {
            if (record[index] != '"') continue;
            if (quoted && index + 1 < record.Length && record[index + 1] == '"') {
                index++;
                continue;
            }

            quoted = !quoted;
        }

        return !quoted;
    }

    private static int CountUnquoted(string text, char delimiter) {
        int count = 0;
        bool quoted = false;
        for (int index = 0; index < text.Length; index++) {
            char ch = text[index];
            if (ch == '"') {
                if (quoted && index + 1 < text.Length && text[index + 1] == '"') index++;
                else quoted = !quoted;
            } else if (ch == delimiter && !quoted) {
                count++;
            }
        }

        return count;
    }
}
