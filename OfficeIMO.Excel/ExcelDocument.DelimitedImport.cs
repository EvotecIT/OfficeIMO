using System.Data;
using System.Globalization;
using OfficeIMO.CSV;

#pragma warning disable CS1591
namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Imports normalized CSV/TSV content into a worksheet using OfficeIMO's tabular writer.
        /// </summary>
        public ExcelDelimitedImportResult ImportDelimitedText(string text, ExcelDelimitedImportOptions? options = null) {
            if (text == null) throw new ArgumentNullException(nameof(text));
            options ??= new ExcelDelimitedImportOptions();
            char delimiter = options.Delimiter ?? DetectDelimitedImportDelimiter(text);
            var warnings = new List<string>();
            using var reader = new StringReader(text);
            DataTable table = ParseDelimitedText(reader, delimiter, options, warnings);
            return ImportDelimitedTable(table, delimiter, "UTF-16 string", options, warnings);
        }

        /// <summary>
        /// Imports normalized CSV/TSV content from a file into a worksheet using OfficeIMO's tabular writer.
        /// </summary>
        public ExcelDelimitedImportResult ImportDelimitedFile(string path, ExcelDelimitedImportOptions? options = null) {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
            options ??= new ExcelDelimitedImportOptions();
            char delimiter = options.Delimiter ?? DetectDelimitedImportDelimiterFromFile(path);
            var warnings = new List<string>();
            DataTable table = ParseDelimitedFile(path, delimiter, options, warnings);
            return ImportDelimitedTable(table, delimiter, "UTF-8 file", options, warnings);
        }

        private ExcelDelimitedImportResult ImportDelimitedTable(DataTable table, char delimiter, string encodingName, ExcelDelimitedImportOptions options, IReadOnlyList<string> warnings) {
            table.TableName = string.IsNullOrWhiteSpace(options.SheetName) ? "Import" : options.SheetName!.Trim();

            var dataSet = new DataSet();
            dataSet.Tables.Add(table);
            IReadOnlyList<ExcelDataSetImportResult> results = InsertDataSet(dataSet, createTables: false, tableStyle: options.TableStyle, includeHeaders: true, includeAutoFilter: true, autoFit: false);
            ExcelDataSetImportResult result = results[0];
            if (options.CreateTable && !string.IsNullOrWhiteSpace(result.Range)) {
                ExcelSheet sheet = this[result.SheetName];
                string requestedTableName = string.IsNullOrWhiteSpace(options.TableName) ? result.SheetName : options.TableName!.Trim();
                sheet.AddTable(result.Range, hasHeader: true, requestedTableName, options.TableStyle, includeAutoFilter: true);
                string? actualTableName = sheet.WorksheetPart.TableDefinitionParts
                    .Select(part => part.Table?.Name?.Value ?? part.Table?.DisplayName?.Value)
                    .FirstOrDefault(name => !string.IsNullOrWhiteSpace(name));
                result = new ExcelDataSetImportResult(result.SheetName, actualTableName, result.Range, result.RowCount, result.ColumnCount);
            }

            return new ExcelDelimitedImportResult(result, delimiter, encodingName, warnings);
        }

        private static DataTable ParseDelimitedFile(string path, char delimiter, ExcelDelimitedImportOptions options, ICollection<string> warnings) {
            var csvOptions = CreateDelimitedCsvLoadOptions(delimiter);
            var table = new DataTable { Locale = options.Culture };
            var recordIndex = 0;

            CsvDocument.ReadRecordsReusable(path, record => {
                AddDelimitedRecord(table, record, recordIndex, options, warnings);
                recordIndex++;
            }, csvOptions);

            return table;
        }

        private static DataTable ParseDelimitedText(TextReader reader, char delimiter, ExcelDelimitedImportOptions options, ICollection<string> warnings) {
            var csvOptions = CreateDelimitedCsvLoadOptions(delimiter);
            var table = new DataTable { Locale = options.Culture };
            var recordIndex = 0;

            CsvDocument.ReadRecordsReusable(reader, record => {
                AddDelimitedRecord(table, record, recordIndex, options, warnings);
                recordIndex++;
            }, csvOptions);

            return table;
        }

        private static CsvLoadOptions CreateDelimitedCsvLoadOptions(char delimiter) =>
            new CsvLoadOptions {
                Delimiter = delimiter,
                HasHeaderRow = false,
                SkipCommentRowsBeforeHeader = false,
                SkipCommentRows = false,
                RecognizeW3CFieldsHeader = false,
                GenerateMissingHeaderNames = false,
                ColumnCountMismatchPolicy = CsvColumnCountMismatchPolicy.PadMissingFieldsAndIgnoreExtraFields
            };

        private static void AddDelimitedRecord(DataTable table, IReadOnlyList<string> record, int recordIndex, ExcelDelimitedImportOptions options, ICollection<string> warnings) {
            if (recordIndex == 0 && options.HeadersInFirstRow) {
                EnsureDelimitedColumns(table, record, useHeaderValues: true);
                return;
            }

            EnsureDelimitedColumns(table, record, useHeaderValues: false);
            DataRow row = table.NewRow();
            for (var columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++) {
                string value = columnIndex < record.Count ? record[columnIndex] : string.Empty;
                row[columnIndex] = ConvertDelimitedValue(value, options, warnings);
            }

            table.Rows.Add(row);
        }

        private static void EnsureDelimitedColumns(DataTable table, IReadOnlyList<string> record, bool useHeaderValues) {
            for (var columnIndex = table.Columns.Count; columnIndex < record.Count; columnIndex++) {
                string name = useHeaderValues && !string.IsNullOrWhiteSpace(record[columnIndex])
                    ? record[columnIndex]
                    : "Column" + (columnIndex + 1).ToString(CultureInfo.InvariantCulture);
                table.Columns.Add(GetUniqueDelimitedColumnName(table, name, columnIndex), typeof(object));
            }
        }

        private static string GetUniqueDelimitedColumnName(DataTable table, string name, int columnIndex) {
            if (!table.Columns.Contains(name)) {
                return name;
            }

            string baseName = name;
            string suffix = "_" + (columnIndex + 1).ToString(CultureInfo.InvariantCulture);
            name = baseName + suffix;
            var duplicateIndex = 2;
            while (table.Columns.Contains(name)) {
                name = baseName + suffix + "_" + duplicateIndex.ToString(CultureInfo.InvariantCulture);
                duplicateIndex++;
            }

            return name;
        }

        private static object ConvertDelimitedValue(string value, ExcelDelimitedImportOptions options, ICollection<string> warnings) {
            if (value.Length == 0) return DBNull.Value;
            if (!options.ConvertNumbersAndDates) return value;
            if (decimal.TryParse(value, NumberStyles.Number, options.Culture, out decimal number)) return number;
            if (DateTime.TryParse(value, options.Culture, DateTimeStyles.None, out DateTime date)) return date;
            return value;
        }

        private static char DetectDelimitedImportDelimiterFromFile(string path) {
            using var reader = new StreamReader(path);
            string firstLine = reader.ReadLine() ?? string.Empty;
            return DetectDelimitedImportDelimiter(firstLine);
        }

        private static char DetectDelimitedImportDelimiter(string text) {
            string firstLine = ReadFirstDelimitedImportLine(text);
            var candidates = new[] { ',', ';', '\t', '|' };
            return candidates
                .Select(candidate => new { Delimiter = candidate, Count = CountUnquoted(firstLine, candidate) })
                .OrderByDescending(item => item.Count)
                .First().Delimiter;
        }

        private static string ReadFirstDelimitedImportLine(string text) {
            var lineEnd = text.IndexOf('\n');
            if (lineEnd < 0) {
                return text;
            }

            if (lineEnd > 0 && text[lineEnd - 1] == '\r') {
                lineEnd--;
            }

            return text.Substring(0, lineEnd);
        }

        private static int CountUnquoted(string text, char delimiter) {
            int count = 0;
            bool quoted = false;
            for (int i = 0; i < text.Length; i++) {
                char ch = text[i];
                if (ch == '"') {
                    if (quoted && i + 1 < text.Length && text[i + 1] == '"') {
                        i++;
                    } else {
                        quoted = !quoted;
                    }
                } else if (ch == delimiter && !quoted) {
                    count++;
                }
            }

            return count;
        }
    }
}
#pragma warning restore CS1591
