using System.Data;
using System.Globalization;
using OfficeIMO.CSV;

namespace OfficeIMO.Reader.Excel;

internal static class ExcelCsvDataTableBuilder {
    internal static DataTable FromText(
        string text,
        char delimiter,
        bool headersInFirstRow,
        int skipInitialRecords,
        CultureInfo culture,
        bool convertNumbersAndDates) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        using var reader = new StringReader(text);
        return FromReader(reader, delimiter, headersInFirstRow, skipInitialRecords, culture, convertNumbersAndDates);
    }

    internal static DataTable FromFile(
        string path,
        char delimiter,
        bool headersInFirstRow,
        int skipInitialRecords,
        CultureInfo culture,
        bool convertNumbersAndDates) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        using var reader = CsvFile.OpenTextReader(path);
        return FromReader(reader, delimiter, headersInFirstRow, skipInitialRecords, culture, convertNumbersAndDates);
    }

    private static DataTable FromReader(
        TextReader reader,
        char delimiter,
        bool headersInFirstRow,
        int skipInitialRecords,
        CultureInfo culture,
        bool convertNumbersAndDates) {
        if (skipInitialRecords < 0) {
            throw new ArgumentOutOfRangeException(nameof(skipInitialRecords), "SkipInitialRecords cannot be negative.");
        }

        var table = new DataTable { Locale = culture ?? CultureInfo.InvariantCulture };
        var recordIndex = 0;
        CsvDocument.ReadRecordsReusable(reader, record => {
            if (skipInitialRecords > 0) {
                skipInitialRecords--;
                return;
            }

            AddRecord(table, record, recordIndex, headersInFirstRow, convertNumbersAndDates);
            recordIndex++;
        }, CreateLoadOptions(delimiter));

        return table;
    }

    private static CsvLoadOptions CreateLoadOptions(char delimiter) =>
        new CsvLoadOptions {
            Delimiter = delimiter,
            HasHeaderRow = false,
            SkipCommentRowsBeforeHeader = false,
            SkipCommentRows = false,
            RecognizeW3CFieldsHeader = false,
            GenerateMissingHeaderNames = false,
            ColumnCountMismatchPolicy = CsvColumnCountMismatchPolicy.PadMissingFieldsAndIgnoreExtraFields
        };

    private static void AddRecord(
        DataTable table,
        IReadOnlyList<string> record,
        int recordIndex,
        bool headersInFirstRow,
        bool convertNumbersAndDates) {
        if (recordIndex == 0 && headersInFirstRow) {
            EnsureColumns(table, record, useHeaderValues: true);
            return;
        }

        EnsureColumns(table, record, useHeaderValues: false);
        DataRow row = table.NewRow();
        for (var columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++) {
            string value = columnIndex < record.Count ? record[columnIndex] : string.Empty;
            row[columnIndex] = ConvertValue(value, table.Locale, convertNumbersAndDates);
        }

        table.Rows.Add(row);
    }

    private static void EnsureColumns(DataTable table, IReadOnlyList<string> record, bool useHeaderValues) {
        for (var columnIndex = table.Columns.Count; columnIndex < record.Count; columnIndex++) {
            string name = useHeaderValues && !string.IsNullOrWhiteSpace(record[columnIndex])
                ? record[columnIndex]
                : "Column" + (columnIndex + 1).ToString(CultureInfo.InvariantCulture);
            table.Columns.Add(GetUniqueColumnName(table, name, columnIndex), typeof(object));
        }
    }

    private static string GetUniqueColumnName(DataTable table, string name, int columnIndex) {
        if (!table.Columns.Contains(name)) return name;

        string suffix = "_" + (columnIndex + 1).ToString(CultureInfo.InvariantCulture);
        string candidate = name + suffix;
        var duplicateIndex = 2;
        while (table.Columns.Contains(candidate)) {
            candidate = name + suffix + "_" + duplicateIndex.ToString(CultureInfo.InvariantCulture);
            duplicateIndex++;
        }

        return candidate;
    }

    private static object ConvertValue(string value, CultureInfo culture, bool convertNumbersAndDates) {
        if (value.Length == 0) return DBNull.Value;
        if (!convertNumbersAndDates) return value;
        if (decimal.TryParse(value, NumberStyles.Number, culture, out decimal number)) return number;
        if (DateTime.TryParse(value, culture, DateTimeStyles.None, out DateTime date)) return date;
        return value;
    }
}
