#if OFFICEIMO_BENCHMARK_NEW_APIS
using ExcelReader.Core.Reader;
using Sylvan.Data.Excel;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>Consumes all native field values after full ciphertext authentication.</summary>
    internal sealed class EncryptedReadWorkload {
        internal EncryptedWorkbookFixture Fixture { get; private set; } = null!;
        private long _checksum;

        internal async Task SetupAsync(EncryptedWorkbookFixture fixture) {
            BenchmarkInput.WriteDescription();
            Fixture = fixture;
            using MemoryStream oracleStream = new MemoryStream(fixture.Plain, writable: false);
            using Sylvan.Data.Excel.ExcelDataReader oracle = global::Sylvan.Data.Excel.ExcelDataReader.Create(oracleStream, ExcelWorkbookType.ExcelXml,
                new ExcelDataReaderOptions { Schema = ExcelSchema.NoHeaders });
            using Stream peerStream = fixture.OpenEncryptedStream();
            using IExcelWorkbook peer = ExcelReaderApi.Open(peerStream, leaveOpen: true, EncryptedWorkbookFixture.PeerOptions());
            using IExcelRowEnumerator rows = peer.FirstSheet.GetEnumerator();
            using ExcelWorkbookDataReader office = ExcelDocument.OpenEncryptedDataReader(fixture.Encrypted, EncryptedWorkbookFixture.Password,
                EncryptedWorkbookFixture.OfficeOptions());
            int count = 0;
            long sum = 0;
            while (oracle.Read()) {
                if (!rows.MoveNext() || !office.Read() || oracle.FieldCount != fixture.Columns
                    || rows.Current.ColumnCount != fixture.Columns || office.FieldCount != fixture.Columns)
                    throw new InvalidDataException("Authenticated reader shape or row count differs from plaintext.");
                TypedRecord? expected = fixture.Columns == 4 ? TypedWorkbookFixture.ExpectedRecord(count + 1) : null;
                for (int column = 0; column < fixture.Columns; column++) {
                    object? value = OracleValue(oracle, column);
                    object? peerValue = PeerValue(rows.Current[column]);
                    object? officeValue = office.IsDBNull(column) ? null : office.GetValue(column);
                    object? generated = expected == null ? value : column switch {
                        0 => expected.Name, 1 => (double)expected.Id, 2 => expected.Date, 3 => expected.Value,
                        _ => throw new InvalidDataException(),
                    };
                    if (!Equals(value, peerValue) || !Equals(value, officeValue) || !Equals(value, generated))
                        throw new InvalidDataException($"Authenticated field/type/order differs at row {count + 1}, column {column + 1}: "
                            + $"oracle={value} ({value?.GetType().Name}), peer={peerValue} ({peerValue?.GetType().Name}), "
                            + $"OfficeIMO={officeValue} ({officeValue?.GetType().Name}).");
                    sum = unchecked(sum + Accumulate(value));
                }
                count++;
            }
            if (count != fixture.Rows || rows.MoveNext() || office.Read() || peer.SheetCount != 1
                || oracle.NextResult() || office.NextResult()) throw new InvalidDataException("Authenticated sheet/row count differs.");
            _checksum = sum;
            ValidateMemoryPeerAndOfficeStream();
            Peer(memory: true);
            Office(memory: true);
            // A prepared generated cache has no files yet. Its measurement stream is memory-backed.
            Peer(memory: false);
            Office(memory: false);
            await ValidateAsyncReaders();
            Console.WriteLine($"Qualified authenticated fields: rows={count}, columns={fixture.Columns}, checksum={sum}; "
                + "every field/type/order matches independent plaintext reader; both engines verify full integrity.");
        }

        internal long Peer(bool memory) {
            using Stream? stream = memory ? null : Fixture.OpenEncryptedStream();
            using IExcelWorkbook workbook = memory
                ? ExcelReaderApi.Open(Fixture.Encrypted.AsMemory(), EncryptedWorkbookFixture.PeerOptions())
                : ExcelReaderApi.Open(stream!, leaveOpen: true, EncryptedWorkbookFixture.PeerOptions());
            long sum = 0;
            int count = 0;
            using IExcelRowEnumerator rows = workbook.FirstSheet.GetEnumerator();
            while (rows.MoveNext()) {
                sum = unchecked(sum + AccumulatePeerRow(rows.Current));
                count++;
            }
            return Check(sum, count);
        }

        internal long Office(bool memory) {
            using Stream? stream = memory ? null : Fixture.OpenEncryptedStream();
            using ExcelWorkbookDataReader reader = memory
                ? ExcelDocument.OpenEncryptedDataReader(Fixture.Encrypted, EncryptedWorkbookFixture.Password, EncryptedWorkbookFixture.OfficeOptions())
                : ExcelDocument.OpenEncryptedDataReader(stream!, EncryptedWorkbookFixture.Password, EncryptedWorkbookFixture.OfficeOptions());
            long sum = 0;
            int count = 0;
            while (reader.Read()) {
                for (int column = 0; column < reader.FieldCount; column++)
                    sum = unchecked(sum + Accumulate(reader.IsDBNull(column) ? null : reader.GetValue(column)));
                count++;
            }
            return Check(sum, count);
        }

        internal async Task<long> PeerAsync() {
            await using Stream stream = Fixture.OpenEncryptedStream();
            await using IExcelWorkbook workbook = await ExcelReaderApi.OpenAsync(stream, leaveOpen: true,
                EncryptedWorkbookFixture.PeerOptions());
            long sum = 0;
            int count = 0;
            await using IExcelRowEnumerator rows = workbook.FirstSheet.GetAsyncEnumerator();
            while (await rows.MoveNextAsync()) {
                sum = unchecked(sum + AccumulatePeerRow(rows.Current));
                count++;
            }
            return Check(sum, count);
        }

        internal async Task<long> OfficeAsync() {
            await using Stream stream = Fixture.OpenEncryptedStream();
            await using ExcelWorkbookDataReader reader = await ExcelDocument.OpenEncryptedDataReaderAsync(stream,
                EncryptedWorkbookFixture.Password, EncryptedWorkbookFixture.OfficeOptions());
            long sum = 0;
            int count = 0;
            while (await reader.ReadAsync()) {
                for (int column = 0; column < reader.FieldCount; column++)
                    sum = unchecked(sum + Accumulate(reader.IsDBNull(column) ? null : reader.GetValue(column)));
                count++;
            }
            return Check(sum, count);
        }

        private async Task ValidateAsyncReaders() {
            await using Stream peerStream = Fixture.OpenEncryptedStream();
            await using IExcelWorkbook peer = await ExcelReaderApi.OpenAsync(peerStream, leaveOpen: true,
                EncryptedWorkbookFixture.PeerOptions());
            await using IExcelRowEnumerator rows = peer.FirstSheet.GetAsyncEnumerator();
            await using Stream officeStream = Fixture.OpenEncryptedStream();
            await using ExcelWorkbookDataReader office = await ExcelDocument.OpenEncryptedDataReaderAsync(officeStream,
                EncryptedWorkbookFixture.Password, EncryptedWorkbookFixture.OfficeOptions());
            using MemoryStream oracleStream = new MemoryStream(Fixture.Plain, writable: false);
            using Sylvan.Data.Excel.ExcelDataReader oracle = global::Sylvan.Data.Excel.ExcelDataReader.Create(oracleStream, ExcelWorkbookType.ExcelXml,
                new ExcelDataReaderOptions { Schema = ExcelSchema.NoHeaders });
            int count = 0;
            while (oracle.Read()) {
                if (!await rows.MoveNextAsync() || !await office.ReadAsync()
                    || rows.Current.ColumnCount != Fixture.Columns || office.FieldCount != Fixture.Columns)
                    throw new InvalidDataException("Async authenticated row shape differs.");
                for (int column = 0; column < Fixture.Columns; column++) {
                    object? value = OracleValue(oracle, column);
                    if (!Equals(value, PeerValue(rows.Current[column]))
                        || !Equals(value, office.IsDBNull(column) ? null : office.GetValue(column)))
                        throw new InvalidDataException($"Async authenticated field differs at {count + 1}/{column + 1}.");
                }
                count++;
            }
            if (count != Fixture.Rows || await rows.MoveNextAsync() || await office.ReadAsync())
                throw new InvalidDataException("Async authenticated row count differs.");
            await PeerAsync();
            await OfficeAsync();
        }

        private void ValidateMemoryPeerAndOfficeStream() {
            using MemoryStream oracleStream = new MemoryStream(Fixture.Plain, writable: false);
            using Sylvan.Data.Excel.ExcelDataReader oracle = global::Sylvan.Data.Excel.ExcelDataReader.Create(oracleStream, ExcelWorkbookType.ExcelXml,
                new ExcelDataReaderOptions { Schema = ExcelSchema.NoHeaders });
            using IExcelWorkbook peer = ExcelReaderApi.Open(Fixture.Encrypted.AsMemory(), EncryptedWorkbookFixture.PeerOptions());
            using IExcelRowEnumerator rows = peer.FirstSheet.GetEnumerator();
            using Stream stream = Fixture.OpenEncryptedStream();
            using ExcelWorkbookDataReader office = ExcelDocument.OpenEncryptedDataReader(stream, EncryptedWorkbookFixture.Password,
                EncryptedWorkbookFixture.OfficeOptions());
            int count = 0;
            while (oracle.Read()) {
                if (!rows.MoveNext() || !office.Read() || rows.Current.ColumnCount != Fixture.Columns
                    || office.FieldCount != Fixture.Columns) throw new InvalidDataException("Memory/stream authenticated shape differs.");
                for (int column = 0; column < Fixture.Columns; column++) {
                    object? value = OracleValue(oracle, column);
                    if (!Equals(value, PeerValue(rows.Current[column]))
                        || !Equals(value, office.IsDBNull(column) ? null : office.GetValue(column)))
                        throw new InvalidDataException($"Memory/stream authenticated field differs at {count + 1}/{column + 1}.");
                }
                count++;
            }
            if (count != Fixture.Rows || rows.MoveNext() || office.Read() || peer.SheetCount != 1 || office.NextResult())
                throw new InvalidDataException("Memory/stream authenticated row/sheet count differs.");
        }

        private long Check(long sum, int rows) => sum == _checksum && rows == Fixture.Rows
            ? sum : throw new InvalidDataException("Authenticated scan checksum/count differs.");
        private static long Accumulate(object? value) => value switch {
            string text => text.Length, double number => (long)number, DateTime date => date.Ticks, _ => 0,
        };
        private static long AccumulatePeerRow(Row row) {
            long sum = 0;
            for (int column = 0; column < row.ColumnCount; column++) {
                Cell cell = row[column];
                switch (cell.Type) {
                    case CellType.Empty: break;
                    case CellType.ExcelString: sum = unchecked(sum + cell.GetString().Length); break;
                    case CellType.Number when cell.TryParse(null, out double number): sum = unchecked(sum + (long)number); break;
                    case CellType.Date when cell.TryGetDateTime(out DateTime date): sum = unchecked(sum + date.Ticks); break;
                    default: throw new InvalidDataException($"Unexpected encrypted field type: {cell.Type}");
                }
            }
            return sum;
        }
        private static object? PeerValue(Cell cell) => cell.Type switch {
            CellType.Empty => null, CellType.ExcelString => cell.GetString(),
            CellType.Number when cell.TryParse(null, out double number) => number,
            CellType.Date when cell.TryGetDateTime(out DateTime date) => date,
            _ => throw new InvalidDataException($"Unexpected pinned encrypted field type: {cell.Type}"),
        };
        private static object? OracleValue(global::Sylvan.Data.Excel.ExcelDataReader reader, int column) {
            if (reader.IsDBNull(column)) return null;
            return reader.GetExcelDataType(column) switch {
                ExcelDataType.String => reader.GetString(column),
                ExcelDataType.Numeric when reader.GetFormat(column)?.Kind == FormatKind.Date => reader.GetDateTime(column),
                ExcelDataType.Numeric => reader.GetDouble(column), ExcelDataType.DateTime => reader.GetDateTime(column),
                _ => throw new InvalidDataException("Unexpected pinned encrypted plaintext field type."),
            };
        }
    }
}
#endif