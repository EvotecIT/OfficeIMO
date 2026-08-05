using System.Data;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Writes strongly typed rows directly into an XLSX package through a row writer.
        /// Sources are consumed once while worksheet XML is written unless table creation requires
        /// the final row range before package output begins.
        /// </summary>
        public static ExcelDataSetImportResult WriteRows<T>(
            Stream stream,
            IEnumerable<T> items,
            IReadOnlyList<string> headers,
            Action<ExcelTabularRowWriter, T> writeRow,
            ExcelTabularWriteOptions? options = null,
            CancellationToken ct = default) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanWrite) throw new ArgumentException("The destination stream must be writable.", nameof(stream));
            if (items == null) throw new ArgumentNullException(nameof(items));
            if (headers == null) throw new ArgumentNullException(nameof(headers));
            if (headers.Count == 0) throw new ArgumentException("At least one column header is required.", nameof(headers));
            if (writeRow == null) throw new ArgumentNullException(nameof(writeRow));

            options = CreateRowWriteOptions(options);
            if (options.AutoFit) {
                throw new ArgumentException("Row-writer exports cannot calculate column widths before writing rows.", nameof(options));
            }

            ValidateTabularHeaders(headers);

            DirectDataSetTableModel tableModel;
            if (CanWriteRowsSinglePass(options)) {
                tableModel = DirectDataSetTableModel.FromStreamingCallbackRows(
                    headers,
                    items,
                    writeRow,
                    GetMaximumDataRows(options));
            } else {
                var rows = MaterializeRows(items, GetMaximumDataRows(options), ct);
                tableModel = DirectDataSetTableModel.FromCallbackRows(headers, rows, writeRow);
            }
            return WriteTabularModel(stream, tableModel, options, ct);
        }

        /// <summary>
        /// Writes an asynchronous sequence of strongly typed rows directly into an XLSX package.
        /// Rows are awaited and written one at a time without buffering the source sequence.
        /// </summary>
        /// <remarks>
        /// Asynchronous row sources require package-native streaming output. Table creation and
        /// automatic column sizing are therefore not supported by this overload.
        /// </remarks>
        public static async Task<ExcelDataSetImportResult> WriteRowsAsync<T>(
            Stream stream,
            IAsyncEnumerable<T> items,
            IReadOnlyList<string> headers,
            Action<ExcelTabularRowWriter, T> writeRow,
            ExcelTabularWriteOptions? options = null,
            CancellationToken ct = default) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanWrite) throw new ArgumentException("The destination stream must be writable.", nameof(stream));
            if (items == null) throw new ArgumentNullException(nameof(items));
            if (headers == null) throw new ArgumentNullException(nameof(headers));
            if (headers.Count == 0) throw new ArgumentException("At least one column header is required.", nameof(headers));
            if (writeRow == null) throw new ArgumentNullException(nameof(writeRow));

            options = CreateRowWriteOptions(options);
            if (options.CreateTable) {
                throw new ArgumentException(
                    "Asynchronous row exports cannot create a table because the final row range is not known before package output begins.",
                    nameof(options));
            }
            if (options.AutoFit) {
                throw new ArgumentException("Asynchronous row exports cannot calculate column widths before writing rows.", nameof(options));
            }

            ValidateTabularHeaders(headers);
            ct.ThrowIfCancellationRequested();

            var columnTypes = new Type[headers.Count];
            for (int columnIndex = 0; columnIndex < columnTypes.Length; columnIndex++) {
                columnTypes[columnIndex] = typeof(object);
            }

            var tableModel = DirectDataSetTableModel.FromRows(headers, columnTypes, Array.Empty<object?[]>());
            string sheetName = DirectDataSetWorkbookModel.SanitizeSheetName(options.SheetName);
            string initialRange = ExcelSheet.BuildObjectExportRange(1, tableModel.ColumnCount, 0, options.IncludeHeaders);
            var model = DirectDataSetWorkbookModel.CreateSingle(
                sheetName,
                sheetName,
                tableName: null,
                initialRange,
                tableModel,
                createTable: false,
                options.TableStyle,
                options.IncludeHeaders,
                options.IncludeAutoFilter,
                autoFit: false,
                DefaultDateTimeOffsetWriteStrategy,
                ct,
                options.UseCellValueNumberFormats,
                options.DateSystem,
                options.IncludeCellReferences);

            if (stream.CanSeek) {
                PrepareDestinationStreamForWrite(stream);
            }

            int rowCount = await DirectDataSetWorkbookWriter.WriteRowsAsync(
                stream,
                model,
                items,
                writeRow,
                GetMaximumDataRows(options),
                ct).ConfigureAwait(false);
            if (stream.CanSeek) {
                stream.Seek(0, SeekOrigin.Begin);
            }

            string range = ExcelSheet.BuildObjectExportRange(
                1,
                tableModel.ColumnCount,
                rowCount,
                options.IncludeHeaders);
            return new ExcelDataSetImportResult(sheetName, tableName: null, range, rowCount, tableModel.ColumnCount);
        }

        private static ExcelTabularWriteOptions CreateRowWriteOptions(ExcelTabularWriteOptions? options) {
            if (options == null) {
                return new ExcelTabularWriteOptions { UseSharedStrings = false };
            }

            return new ExcelTabularWriteOptions {
                SheetName = options.SheetName,
                IncludeHeaders = options.IncludeHeaders,
                CreateTable = options.CreateTable,
                TableName = options.TableName,
                TableStyle = options.TableStyle,
                IncludeAutoFilter = options.IncludeAutoFilter,
                AutoFit = options.AutoFit,
                UseCellValueNumberFormats = options.UseCellValueNumberFormats,
                IncludeCellReferences = options.IncludeCellReferences,
                UseSharedStrings = false,
                DateSystem = options.DateSystem
            };
        }

        private static bool CanWriteRowsSinglePass(ExcelTabularWriteOptions options) =>
            !options.UseSharedStrings
            && !options.CreateTable
            && !options.AutoFit;

        private static void ValidateTabularHeaders(IReadOnlyList<string> headers) {
            var usedHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int columnIndex = 0; columnIndex < headers.Count; columnIndex++) {
                string header = headers[columnIndex] ?? string.Empty;
                if (string.IsNullOrWhiteSpace(header)) {
                    throw new ArgumentException("Column headers must not be empty.", nameof(headers));
                }
                if (!usedHeaders.Add(header)) {
                    throw new ArgumentException("Column headers must be unique.", nameof(headers));
                }
            }
        }

        private static int GetMaximumDataRows(ExcelTabularWriteOptions options) =>
            A1.MaxRows - (options.IncludeHeaders ? 1 : 0);

        private static IReadOnlyList<T> MaterializeRows<T>(
            IEnumerable<T> items,
            int maximumDataRows,
            CancellationToken ct) {
            ct.ThrowIfCancellationRequested();
            if (items is IReadOnlyList<T> rows) {
                if (rows.Count > maximumDataRows) {
                    throw new InvalidOperationException("Row export exceeds the maximum worksheet row count.");
                }

                return rows;
            }

            var materialized = new List<T>();
            using IEnumerator<T> enumerator = items.GetEnumerator();
            while (true) {
                ct.ThrowIfCancellationRequested();
                if (!enumerator.MoveNext()) {
                    return materialized;
                }

                ct.ThrowIfCancellationRequested();
                if (materialized.Count >= maximumDataRows) {
                    throw new InvalidOperationException("Row export exceeds the maximum worksheet row count.");
                }

                materialized.Add(enumerator.Current);
            }
        }

        /// <summary>
        /// Writes an open data reader directly to an XLSX package without creating an editable workbook object.
        /// The caller owns the reader and its connection.
        /// </summary>
        public static ExcelDataSetImportResult WriteDataReader(
            Stream stream,
            IDataReader reader,
            ExcelTabularWriteOptions? options = null,
            CancellationToken ct = default) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanWrite) throw new ArgumentException("The destination stream must be writable.", nameof(stream));
            if (reader == null) throw new ArgumentNullException(nameof(reader));
            if (reader.FieldCount < 1) throw new ArgumentException("Data reader must expose at least one field.", nameof(reader));

            options ??= new ExcelTabularWriteOptions();
            string[] headers = ExcelSheet.BuildReaderHeaders(reader);
            Type[] fieldTypes = ExcelSheet.BuildReaderFieldTypes(reader);
            string[] columnNames = ExcelSheet.BuildDirectReaderColumnNames(headers, options.IncludeHeaders);
            Type[] columnTypes = ExcelSheet.BuildDirectReaderColumnTypes(fieldTypes);
            if (!options.UseSharedStrings && !options.CreateTable && !options.AutoFit) {
                return WriteStreamingDataReader(stream, reader, columnNames, columnTypes, options, ct);
            }

            var rows = new List<object?[]>(4096);
            bool canCancel = ct.CanBeCanceled;
            bool useBulkRead = ExcelSheet.CanUseBulkDataReaderValues(reader);
            int maximumDataRows = A1.MaxRows - (options.IncludeHeaders ? 1 : 0);
            while (reader.Read()) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                if (rows.Count >= maximumDataRows) {
                    throw new InvalidOperationException("Data reader export exceeds the maximum worksheet row count.");
                }

                var values = new object?[columnNames.Length];
                ExcelSheet.FillDataReaderValues(reader, values, columnNames.Length, ref useBulkRead);
                rows.Add(values);
            }

            var tableModel = DirectDataSetTableModel.FromRows(columnNames, columnTypes, rows);
            return WriteTabularModel(stream, tableModel, options, ct);
        }

        private static ExcelDataSetImportResult WriteStreamingDataReader(
            Stream stream,
            IDataReader reader,
            string[] columnNames,
            Type[] columnTypes,
            ExcelTabularWriteOptions options,
            CancellationToken ct) {
            var tableModel = DirectDataSetTableModel.FromRows(columnNames, columnTypes, Array.Empty<object?[]>());
            string sheetName = DirectDataSetWorkbookModel.SanitizeSheetName(options.SheetName);
            string initialRange = ExcelSheet.BuildObjectExportRange(1, tableModel.ColumnCount, 0, options.IncludeHeaders);
            var model = DirectDataSetWorkbookModel.CreateSingle(
                sheetName,
                sheetName,
                tableName: null,
                initialRange,
                tableModel,
                createTable: false,
                options.TableStyle,
                options.IncludeHeaders,
                options.IncludeAutoFilter,
                autoFit: false,
                DefaultDateTimeOffsetWriteStrategy,
                ct,
                options.UseCellValueNumberFormats,
                options.DateSystem,
                options.IncludeCellReferences);

            if (stream.CanSeek) {
                PrepareDestinationStreamForWrite(stream);
            }

            int rowCount = DirectDataSetWorkbookWriter.WriteDataReader(stream, model, reader, ct);
            if (stream.CanSeek) {
                stream.Seek(0, SeekOrigin.Begin);
            }

            string range = ExcelSheet.BuildObjectExportRange(1, tableModel.ColumnCount, rowCount, options.IncludeHeaders);
            return new ExcelDataSetImportResult(sheetName, tableName: null, range, rowCount, tableModel.ColumnCount);
        }

        private static ExcelDataSetImportResult WriteTabularModel(
            Stream stream,
            DirectDataSetTableModel tableModel,
            ExcelTabularWriteOptions? options,
            CancellationToken ct) {
            if (!stream.CanWrite) throw new ArgumentException("The destination stream must be writable.", nameof(stream));
            options ??= new ExcelTabularWriteOptions();
            string sheetName = DirectDataSetWorkbookModel.SanitizeSheetName(options.SheetName);
            string range = ExcelSheet.BuildObjectExportRange(1, tableModel.ColumnCount, tableModel.RowCount, options.IncludeHeaders);
            var model = DirectDataSetWorkbookModel.CreateSingle(
                sheetName,
                sheetName,
                options.TableName,
                range,
                tableModel,
                options.CreateTable,
                options.TableStyle,
                options.IncludeHeaders,
                options.IncludeAutoFilter,
                options.AutoFit,
                DefaultDateTimeOffsetWriteStrategy,
                ct,
                options.UseCellValueNumberFormats,
                options.DateSystem,
                options.IncludeCellReferences);

            if (stream.CanSeek) {
                PrepareDestinationStreamForWrite(stream);
            }

            DirectDataSetWorkbookWriter.Write(stream, model, ct, disableSharedStrings: !options.UseSharedStrings);
            if (stream.CanSeek) {
                stream.Seek(0, SeekOrigin.Begin);
            }

            ExcelDataSetImportResult result = model.Results[0];
            if (result.RowCount == tableModel.RowCount) {
                return result;
            }

            string finalRange = ExcelSheet.BuildObjectExportRange(
                1,
                tableModel.ColumnCount,
                tableModel.RowCount,
                options.IncludeHeaders);
            return new ExcelDataSetImportResult(
                sheetName,
                tableName: null,
                finalRange,
                tableModel.RowCount,
                tableModel.ColumnCount);
        }
    }
}
