using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xls;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Reader.Xlsx;
using Sylvan.Data.Excel;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    public partial class RawReadBenchmarks {
        internal async Task ValidateAsyncReaders() {
            await using (MemoryStream stream = new MemoryStream(_workbook, writable: false)) {
                using IExcelWorkbook workbook = Format switch {
                    RawWorkbookFormat.Xlsx => await ExcelReaderApi.FromXlsxAsync(stream),
                    RawWorkbookFormat.Xlsb => await ExcelReaderApi.FromXlsbAsync(stream),
                    RawWorkbookFormat.Xls => await ExcelReaderApi.FromXlsAsync(stream),
                    _ => throw new ArgumentOutOfRangeException(nameof(Format)),
                };
                int count = 0;
                await using IExcelRowEnumerator enumerator = workbook.FirstSheet.GetAsyncEnumerator();
                while (await enumerator.MoveNextAsync()) ValidateExcelReaderRow(enumerator.Current, ++count);
                if (count != RowCount || workbook.SheetCount != 1) throw new InvalidDataException("Incorrect async raw row/sheet count.");
            }
            await using (ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(_workbook, new ExcelReadOptions { HasHeaderRow = false })) {
                int count = 0;
                while (await reader.ReadAsync()) {
                    if (reader.FieldCount != 4 || reader.GetValue(0) is not string name || reader.GetValue(1) is not double id
                        || reader.GetValue(2) is not DateTime date || reader.GetValue(3) is not double value)
                        throw new InvalidDataException("OfficeIMO returned incorrect async raw types or shape.");
                    TypedWorkbookFixture.ValidateRecord(new TypedRecord { Name = name, Id = (int)id, Date = date, Value = value }, ++count);
                    if (id != count) throw new InvalidDataException("Incorrect async numeric ID.");
                }
                if (count != RowCount || await reader.NextResultAsync()) throw new InvalidDataException("Incorrect async raw row/sheet count.");
            }
            await using (MemoryStream stream = new MemoryStream(_workbook, writable: false)) {
                await using Sylvan.Data.Excel.ExcelDataReader reader = await OpenSylvanAsync(stream);
                int count = 0;
                while (await reader.ReadAsync()) {
                    if (reader.FieldCount != 4 || reader.GetExcelDataType(0) != ExcelDataType.String
                        || reader.GetExcelDataType(1) != ExcelDataType.Numeric || reader.GetExcelDataType(2) != ExcelDataType.Numeric
                        || reader.GetFormat(2)?.Kind != FormatKind.Date || reader.GetExcelDataType(3) != ExcelDataType.Numeric)
                        throw new InvalidDataException("Sylvan returned incorrect async raw types or shape.");
                    double id = reader.GetDouble(1);
                    TypedWorkbookFixture.ValidateRecord(new TypedRecord { Name = reader.GetString(0), Id = (int)id,
                        Date = reader.GetDateTime(2), Value = reader.GetDouble(3) }, ++count);
                    if (id != count) throw new InvalidDataException("Incorrect async numeric ID.");
                }
                if (count != RowCount || await reader.NextResultAsync()) throw new InvalidDataException("Incorrect async raw row/sheet count.");
            }
            await ExcelReaderAsync(materializeStrings: false);
            await ExcelReaderAsync(materializeStrings: true);
            await OfficeIMOAsync();
            await SylvanAsync();
            Console.WriteLine($"Validated async raw {Format}: rows={RowCount}; all fields checked.");
        }

        internal Task<long> ExcelReaderAsync(bool materializeStrings) => Format switch {
            RawWorkbookFormat.Xlsx => ReadXlsxAsync(materializeStrings),
            RawWorkbookFormat.Xlsb => ReadXlsbAsync(materializeStrings),
            RawWorkbookFormat.Xls => ReadXlsAsync(materializeStrings),
            _ => throw new ArgumentOutOfRangeException(nameof(Format)),
        };

        private async Task<long> ReadXlsxAsync(bool materializeStrings) {
            await using MemoryStream stream = new MemoryStream(_workbook, writable: false);
            await using XlsxWorkbook workbook = await ExcelReaderApi.FromXlsxAsync(stream);
            await using XlsxWorkbook.Enumerator enumerator = workbook.FirstSheet.GetAsyncEnumerator();
            long sum = 0;
            int count = 0;
            while (await enumerator.MoveNextAsync()) {
                sum = unchecked(sum + AccumulateRow(enumerator.Current, materializeStrings));
                count++;
            }
            return Check(sum, count);
        }

        private async Task<long> ReadXlsbAsync(bool materializeStrings) {
            await using MemoryStream stream = new MemoryStream(_workbook, writable: false);
            await using XlsbWorkbook workbook = await ExcelReaderApi.FromXlsbAsync(stream);
            await using XlsbWorkbook.Enumerator enumerator = workbook.FirstSheet.GetAsyncEnumerator();
            long sum = 0;
            int count = 0;
            while (await enumerator.MoveNextAsync()) {
                sum = unchecked(sum + AccumulateRow(enumerator.Current, materializeStrings));
                count++;
            }
            return Check(sum, count);
        }

        private async Task<long> ReadXlsAsync(bool materializeStrings) {
            await using MemoryStream stream = new MemoryStream(_workbook, writable: false);
            await using XlsWorkbook workbook = await ExcelReaderApi.FromXlsAsync(stream);
            await using XlsWorkbook.Enumerator enumerator = workbook.FirstSheet.GetAsyncEnumerator();
            long sum = 0;
            int count = 0;
            while (await enumerator.MoveNextAsync()) {
                sum = unchecked(sum + AccumulateRow(enumerator.Current, materializeStrings));
                count++;
            }
            return Check(sum, count);
        }

        internal async Task<long> OfficeIMOAsync() {
            await using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(_workbook, new ExcelReadOptions { HasHeaderRow = false });
            long sum = 0;
            int count = 0;
            while (await reader.ReadAsync()) {
                for (int column = 0; column < reader.FieldCount; column++) {
                    switch (reader.GetValue(column)) {
                        case string value: sum = unchecked(sum + value.Length); break;
                        case double value: sum = unchecked(sum + (long)value); break;
                        case DateTime value: sum = unchecked(sum + value.Ticks); break;
                    }
                }
                count++;
            }
            return Check(sum, count);
        }

        internal async Task<long> SylvanAsync() {
            await using MemoryStream stream = new MemoryStream(_workbook, writable: false);
            await using Sylvan.Data.Excel.ExcelDataReader reader = await OpenSylvanAsync(stream);
            long sum = 0;
            int count = 0;
            do {
                while (await reader.ReadAsync()) {
                    for (int column = 0; column < reader.FieldCount; column++) {
                        if (reader.IsDBNull(column)) continue;
                        switch (reader.GetExcelDataType(column)) {
                            case ExcelDataType.String: sum = unchecked(sum + reader.GetString(column).Length); break;
                            case ExcelDataType.Numeric:
                                sum = unchecked(sum + (reader.GetFormat(column)?.Kind == FormatKind.Date
                                    ? reader.GetDateTime(column).Ticks : (long)reader.GetDouble(column)));
                                break;
                            case ExcelDataType.DateTime: sum = unchecked(sum + reader.GetDateTime(column).Ticks); break;
                        }
                    }
                    count++;
                }
            } while (await reader.NextResultAsync());
            return Check(sum, count);
        }

        private Task<global::Sylvan.Data.Excel.ExcelDataReader> OpenSylvanAsync(Stream stream) =>
            global::Sylvan.Data.Excel.ExcelDataReader.CreateAsync(stream, Format switch {
                RawWorkbookFormat.Xlsx => ExcelWorkbookType.ExcelXml,
                RawWorkbookFormat.Xlsb => ExcelWorkbookType.ExcelBinary,
                RawWorkbookFormat.Xls => ExcelWorkbookType.Excel,
                _ => throw new ArgumentOutOfRangeException(nameof(Format)),
            }, new ExcelDataReaderOptions { Schema = ExcelSchema.NoHeaders });
    }
}
