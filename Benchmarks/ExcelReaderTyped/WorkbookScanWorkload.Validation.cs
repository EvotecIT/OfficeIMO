using ExcelReader.Core.Reader;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    internal sealed partial class WorkbookScanWorkload {
        /// <summary>Qualifies exact first-sheet values and row order against two independent readers.</summary>
        private (long Raw, long Materialized, int Sheets) QualifyValues(
            ExcelReader.Core.Reader.ExcelReaderOptions? options = null, bool borrowedUtf8 = false) {
            using MemoryStream stream = new MemoryStream(_bytes, writable: false);
            using IExcelWorkbook workbook = OpenPeer(stream, options);
            using ExcelWorkbookDataReader office = ExcelDocument.OpenReader(new MemoryStream(_bytes, writable: false),
                new OfficeIMO.Excel.ExcelReaderOptions { HasHeaderRow = false, SheetIndex = 0 });
            using MemoryStream sylvanStream = new MemoryStream(_bytes, writable: false);
            using Sylvan.Data.Excel.ExcelDataReader sylvan = OpenSylvan(sylvanStream);
            int count = 0;
            long raw = 0, materialized = 0;
            foreach (Row row in workbook.FirstSheet) {
                if (row.ColumnCount != _columns || !office.Read() || !sylvan.Read()
                    || office.FieldCount != _columns || sylvan.FieldCount != _columns)
                    throw new InvalidDataException("Reader row count or width differs during full qualification.");
                object[]? expected = _expectedValues?.Invoke(count);
                for (int column = 0; column < _columns; column++) {
                    object? value = PeerValue(row[column]);
                    object? officeValue = office.IsDBNull(column) ? null : office.GetValue(column);
                    object? sylvanValue = SylvanValue(sylvan, column);
                    if (!Equals(value, officeValue) || !Equals(value, sylvanValue)
                        || (expected != null && !Equals(value, expected[column])))
                        throw new InvalidDataException($"Field/type mismatch at row {count + 1}, column {column + 1}: "
                            + $"ExcelReader={value} ({value?.GetType().Name}), OfficeIMO={officeValue} ({officeValue?.GetType().Name}), "
                            + $"Sylvan={sylvanValue} ({sylvanValue?.GetType().Name}).");
#if OFFICEIMO_BENCHMARK_NEW_APIS
                    if (borrowedUtf8) QualifyBorrowedText(row[column], office, column, value, count);
#else
                    if (borrowedUtf8) throw new InvalidOperationException("Public UTF8 qualification requires the new API lane.");
#endif
                    raw = unchecked(raw + AccumulateValue(value, byteLength: true));
                    materialized = unchecked(materialized + AccumulateValue(value, byteLength: false));
                }
                count++;
            }
            if (count != _rows || office.Read() || sylvan.Read())
                throw new InvalidDataException("Reader row counts differ.");
            return (raw, materialized, workbook.SheetCount);
        }
    }
}
