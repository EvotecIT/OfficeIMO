using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xls;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Reader.Xlsx;
using Sylvan.Data.Excel;
using System.Data.Common;
using System.Text;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>The formats of the pinned real-data and generated string-heavy read inputs.</summary>
    public enum ComparisonWorkbookFormat { Xlsx, Xlsm, Xlsb, Xls }

    /// <summary>Shares validated scan operations between the two larger input suites.</summary>
    internal sealed class WorkbookScanWorkload {
        private static readonly ExcelReaderOptions PrefetchOptions = new() { PrefetchDecompression = true };
        private static readonly ExcelReaderOptions InternOptions = new() { InternStrings = true };
        private byte[] _bytes = [];
        private ComparisonWorkbookFormat _format;
        private int _rows;
        private long _spanChecksum, _stringChecksum;

        internal void Setup(byte[] bytes, ComparisonWorkbookFormat format, int rows, int columns,
            Func<int, object[]>? expectedValues = null) {
            _bytes = bytes;
            _format = format;
            _rows = rows;
            BenchmarkInput.WriteWorkbookFixtureIdentity($"validated-scan/{format}/rows={rows}/columns={columns}", bytes,
                zipPackage: format != ComparisonWorkbookFormat.Xls);
            using MemoryStream stream = new MemoryStream(bytes, writable: false);
            using IExcelWorkbook workbook = OpenPeer(stream);
            using ExcelWorkbookDataReader office = ExcelDocument.OpenDataReader(bytes, new ExcelReadOptions { HasHeaderRow = false, SheetIndex = 0 });
            using MemoryStream sylvanStream = new MemoryStream(bytes, writable: false);
            using Sylvan.Data.Excel.ExcelDataReader sylvan = OpenSylvan(sylvanStream);
            int count = 0;
            long raw = 0, materialized = 0;
            foreach (Row row in workbook.FirstSheet) {
                if (row.ColumnCount != columns || !office.Read() || !sylvan.Read()
                    || office.FieldCount != columns || sylvan.FieldCount != columns)
                    throw new InvalidDataException("Reader row count or width differs during full qualification.");
                object[]? expected = expectedValues?.Invoke(count);
                for (int column = 0; column < columns; column++) {
                    object? value = PeerValue(row[column]);
                    object? officeValue = office.IsDBNull(column) ? null : office.GetValue(column);
                    object? sylvanValue = SylvanValue(sylvan, column);
                    if (!Equals(value, officeValue) || !Equals(value, sylvanValue)
                        || (expected != null && !Equals(value, expected[column])))
                        throw new InvalidDataException($"Field/type mismatch at row {count + 1}, column {column + 1}: "
                            + $"ExcelReader={value} ({value?.GetType().Name}), OfficeIMO={officeValue} ({officeValue?.GetType().Name}), "
                            + $"Sylvan={sylvanValue} ({sylvanValue?.GetType().Name}).");
                    raw = unchecked(raw + AccumulateValue(value, byteLength: true));
                    materialized = unchecked(materialized + AccumulateValue(value, byteLength: false));
                }
                count++;
            }
            if (count != rows || office.Read() || sylvan.Read()) throw new InvalidDataException("Reader row counts differ.");
            _spanChecksum = raw;
            _stringChecksum = materialized;
            ExcelReader();
            ExcelReader(prefetch: true);
            ExcelReader(memory: true);
            ExcelReader(prefetch: true, memory: true);
            ExcelReader(materializeStrings: true);
            ExcelReader(materializeStrings: true, intern: true);
            OfficeIMO();
            OfficeIMO(streamInput: true);
            Sylvan();
            Console.WriteLine($"Qualified {_format}: rowsRead={rows}, columns={columns}, bytes={bytes.Length}, "
                + $"spanChecksum={raw}, stringChecksum={materialized}, peerSheetCount={workbook.SheetCount}.");
        }

        internal long ExcelReader(bool materializeStrings = false, bool prefetch = false, bool memory = false, bool intern = false) {
            ExcelReaderOptions? options = prefetch ? PrefetchOptions : intern ? InternOptions : null;
            return _format switch {
                ComparisonWorkbookFormat.Xlsx or ComparisonWorkbookFormat.Xlsm => ReadXlsx(materializeStrings, options, memory),
                ComparisonWorkbookFormat.Xlsb => ReadXlsb(materializeStrings, options, memory),
                ComparisonWorkbookFormat.Xls => ReadXls(materializeStrings, options, memory),
                _ => throw new ArgumentOutOfRangeException(),
            };
        }

        private long ReadXlsx(bool materialize, ExcelReaderOptions? options, bool memory) {
            using MemoryStream? stream = memory ? null : new MemoryStream(_bytes, writable: false);
            using XlsxWorkbook workbook = memory ? ExcelReaderApi.FromXlsx(_bytes.AsMemory(), options)
                : ExcelReaderApi.FromXlsx(stream!, options: options);
            long sum = 0;
            int rows = 0;
            foreach (Row row in workbook.FirstSheet) { sum = unchecked(sum + AccumulateRow(row, materialize)); rows++; }
            return Check(sum, rows, materialize ? _stringChecksum : _spanChecksum);
        }

        private long ReadXlsb(bool materialize, ExcelReaderOptions? options, bool memory) {
            using MemoryStream? stream = memory ? null : new MemoryStream(_bytes, writable: false);
            using XlsbWorkbook workbook = memory ? ExcelReaderApi.FromXlsb(_bytes.AsMemory(), options)
                : ExcelReaderApi.FromXlsb(stream!, options: options);
            long sum = 0;
            int rows = 0;
            foreach (Row row in workbook.FirstSheet) { sum = unchecked(sum + AccumulateRow(row, materialize)); rows++; }
            return Check(sum, rows, materialize ? _stringChecksum : _spanChecksum);
        }

        private long ReadXls(bool materialize, ExcelReaderOptions? options, bool memory) {
            using MemoryStream? stream = memory ? null : new MemoryStream(_bytes, writable: false);
            using XlsWorkbook workbook = memory ? ExcelReaderApi.FromXls(_bytes.AsMemory(), options)
                : ExcelReaderApi.FromXls(stream!, options: options);
            long sum = 0;
            int rows = 0;
            foreach (Row row in workbook.FirstSheet) { sum = unchecked(sum + AccumulateRow(row, materialize)); rows++; }
            return Check(sum, rows, materialize ? _stringChecksum : _spanChecksum);
        }

        internal long OfficeIMO(bool streamInput = false) {
            using MemoryStream? stream = streamInput ? new MemoryStream(_bytes, writable: false) : null;
            using ExcelWorkbookDataReader reader = streamInput
                ? ExcelDocument.OpenDataReader(stream!, new ExcelReadOptions { HasHeaderRow = false, SheetIndex = 0 })
                : ExcelDocument.OpenDataReader(_bytes, new ExcelReadOptions { HasHeaderRow = false, SheetIndex = 0 });
            long sum = 0;
            int rows = 0;
            while (reader.Read()) {
                for (int column = 0; column < reader.FieldCount; column++) {
                    switch (reader.GetValue(column)) {
                        case string value: sum = unchecked(sum + value.Length); break;
                        case double value: sum = unchecked(sum + (long)value); break;
                        case DateTime value: sum = unchecked(sum + value.Ticks); break;
                    }
                }
                rows++;
            }
            return Check(sum, rows, _stringChecksum);
        }

        internal long Sylvan() {
            using MemoryStream stream = new MemoryStream(_bytes, writable: false);
            using Sylvan.Data.Excel.ExcelDataReader reader = OpenSylvan(stream);
            long sum = 0;
            int rows = 0;
            while (reader.Read()) {
                for (int column = 0; column < reader.FieldCount; column++) {
                    if (reader.IsDBNull(column)) continue;
                    switch (reader.GetExcelDataType(column)) {
                        case ExcelDataType.String: sum = unchecked(sum + reader.GetString(column).Length); break;
                        case ExcelDataType.Numeric:
                            sum = unchecked(sum + (reader.GetFormat(column)?.Kind == FormatKind.Date
                                ? reader.GetDateTime(column).Ticks : (long)reader.GetDouble(column))); break;
                        case ExcelDataType.DateTime: sum = unchecked(sum + reader.GetDateTime(column).Ticks); break;
                    }
                }
                rows++;
            }
            return Check(sum, rows, _stringChecksum);
        }

        private IExcelWorkbook OpenPeer(Stream stream) => _format switch {
            ComparisonWorkbookFormat.Xlsx or ComparisonWorkbookFormat.Xlsm => ExcelReaderApi.FromXlsx(stream),
            ComparisonWorkbookFormat.Xlsb => ExcelReaderApi.FromXlsb(stream),
            ComparisonWorkbookFormat.Xls => ExcelReaderApi.FromXls(stream),
            _ => throw new ArgumentOutOfRangeException(),
        };

        private global::Sylvan.Data.Excel.ExcelDataReader OpenSylvan(Stream stream) =>
            global::Sylvan.Data.Excel.ExcelDataReader.Create(stream, _format switch {
                ComparisonWorkbookFormat.Xlsx or ComparisonWorkbookFormat.Xlsm => ExcelWorkbookType.ExcelXml,
                ComparisonWorkbookFormat.Xlsb => ExcelWorkbookType.ExcelBinary,
                ComparisonWorkbookFormat.Xls => ExcelWorkbookType.Excel,
                _ => throw new ArgumentOutOfRangeException(),
            }, new ExcelDataReaderOptions { Schema = ExcelSchema.NoHeaders });

        private static object? PeerValue(Cell cell) => cell.Type switch {
            CellType.Empty => null,
            CellType.ExcelString => cell.GetString(),
            CellType.Number when cell.TryParse(null, out double number) => number,
            CellType.Date when cell.TryGetDateTime(out DateTime date) => date,
            _ => throw new InvalidDataException($"Unsupported oracle field type {cell.Type}."),
        };

        private static object? SylvanValue(global::Sylvan.Data.Excel.ExcelDataReader reader, int column) {
            if (reader.IsDBNull(column)) return null;
            return reader.GetExcelDataType(column) switch {
                ExcelDataType.String => reader.GetString(column),
                ExcelDataType.Numeric when reader.GetFormat(column)?.Kind == FormatKind.Date => reader.GetDateTime(column),
                ExcelDataType.Numeric => reader.GetDouble(column),
                ExcelDataType.DateTime => reader.GetDateTime(column),
                _ => throw new InvalidDataException("Unsupported independent reader field type."),
            };
        }

        private static long AccumulateValue(object? value, bool byteLength) => value switch {
            string text => byteLength ? Encoding.UTF8.GetByteCount(text) : text.Length,
            double number => (long)number,
            DateTime date => date.Ticks,
            _ => 0,
        };

        private static long AccumulateRow(Row row, bool materializeStrings) {
            long sum = 0;
            foreach (RowCell rowCell in row.Cells) {
                Cell cell = rowCell.Value;
                switch (cell.Type) {
                    case CellType.ExcelString: sum = unchecked(sum + (materializeStrings ? cell.GetString().Length : cell.Value.Length)); break;
                    case CellType.Number: if (cell.TryParse(null, out double number)) sum = unchecked(sum + (long)number); break;
                    case CellType.Date: if (cell.TryGetDateTime(out DateTime date)) sum = unchecked(sum + date.Ticks); break;
                }
            }
            return sum;
        }

        private long Check(long sum, int rows, long expected) => sum == expected && rows == _rows
            ? sum : throw new InvalidDataException($"Scan count/checksum differs: rows={rows}, sum={sum}, expected={expected}.");
    }
}
