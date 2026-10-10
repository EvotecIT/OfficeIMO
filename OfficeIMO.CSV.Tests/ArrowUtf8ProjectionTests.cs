#if NET8_0_OR_GREATER
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.Common;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Apache.Arrow;
using OfficeIMO.CSV;
using OfficeIMO.Data.Arrow;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public sealed class ArrowUtf8ProjectionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PublicUtf8ProjectionOwnsValuesAndFallsBackPerField(bool asynchronous) {
        object?[][] rows = {
            new object?[] { "Alpha", string.Empty, null, "normalized:one", new Uri("https://example.test/one") },
            new object?[] { "Zażółć 😀", "last", "present", "normalized:two", new Uri("https://example.test/two") },
        };
        using var reader = new BorrowedUtf8Reader(rows);
        var options = new ArrowReadOptions {
            BatchSize = 1,
            ColumnNullability = new[] { false, false, true, false, false },
        };
        List<RecordBatch> batches = await ReadBatches(reader, options, asynchronous);
        try {
            Assert.False(reader.IsClosed);
            // Only the normalized fallback field and unsupported CLR type need
            // ordinary values. A provider may lend empty text for a null field;
            // the reader's DBNull contract still distinguishes it from empty text.
            Assert.Equal(4, reader.MaterializedValueCount);
            Assert.Equal(0, reader.BorrowAttempts[4]);
            reader.Close();
            Assert.Equal(rows.Length, batches.Count);
            for (int row = 0; row < rows.Length; row++) {
                Assert.Equal(1, batches[row].Length);
                for (int column = 0; column < rows[row].Length; column++) {
                    var values = Assert.IsType<StringArray>(batches[row].Column(column));
                    Assert.Equal(rows[row][column] is null, values.IsNull(0));
                    string? expected = rows[row][column] is null
                        ? null : Convert.ToString(rows[row][column], CultureInfo.InvariantCulture);
                    Assert.Equal(expected, values.GetString(0));
                }
            }
        } finally {
            foreach (RecordBatch batch in batches) batch.Dispose();
        }
    }

    [Fact]
    public void ForcedStringColumnsRetainInvariantScalarConversion() {
        object?[][] rows = {
            new object?[] { "Alpha", string.Empty, null, "normalized:one", 1.25m },
            new object?[] { "Beta", "last", "present", "normalized:two", 2.50m },
        };
        using var reader = new BorrowedUtf8Reader(rows, typeof(decimal), CultureInfo.GetCultureInfo("de-DE"));
        using RecordBatch batch = Assert.Single(reader.ReadArrowBatches(new ArrowReadOptions {
            ColumnTypes = Enumerable.Repeat(typeof(string), reader.FieldCount).ToArray(),
        }));

        Assert.Equal(new[] { "1.25", "2.50" }, ReadStrings(batch, 4));
        Assert.Equal(0, reader.BorrowAttempts[4]);
        Assert.Equal(4, reader.MaterializedValueCount);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public async Task CsvStringBatchesPreserveEveryFieldAcrossBufferRefills(bool asynchronous, bool explicitHeader) {
        const int rowCount = 513;
        string repeated = new string('x', 700);
        string[][] rows = Enumerable.Range(0, rowCount).Select(row => new[] {
            row.ToString(CultureInfo.InvariantCulture),
            "Zażółć 😀 " + row.ToString(CultureInfo.InvariantCulture),
            string.Empty,
            repeated + row.ToString(CultureInfo.InvariantCulture),
        }).ToArray();
        string[] header = { "Id", "Name", "Empty", "Long" };
        string csv = (explicitHeader ? string.Empty : "Id,Name,Empty,Long\n")
            + string.Concat(rows.Select(row => string.Join(",", row) + "\n"));
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(csv));
        using var reader = CsvDocument.OpenDataReader(stream, new CsvLoadOptions {
            Header = explicitHeader ? header : null, HasHeaderRow = !explicitHeader,
        });
        Assert.Equal(header, Enumerable.Range(0, reader.FieldCount).Select(reader.GetName).ToArray());
        List<RecordBatch> batches = await ReadBatches(reader, new ArrowReadOptions { BatchSize = 17 }, asynchronous);
        try {
            Assert.False(reader.IsClosed);
            reader.Close();
            Assert.True(stream.CanRead);
            int rowIndex = 0;
            foreach (RecordBatch batch in batches) {
                Assert.InRange(batch.Length, 1, 17);
                for (int row = 0; row < batch.Length; row++, rowIndex++) {
                    for (int column = 0; column < rows[rowIndex].Length; column++) {
                        Assert.Equal(rows[rowIndex][column], Assert.IsType<StringArray>(batch.Column(column)).GetString(row));
                        Assert.False(batch.Column(column).IsNull(row));
                    }
                }
            }
            Assert.Equal(rowCount, rowIndex);
        } finally {
            foreach (RecordBatch batch in batches) batch.Dispose();
        }
    }

    [Theory]
    [InlineData(null)]
    [InlineData("NULL")]
    public void CsvStringProjectionPreservesQuotedNormalizationMissingAndNullMarkers(string? nullValue) {
        const string csv = "Name,Empty,Optional\nplain,,value\n\"a\"\"b\nend\",\"\",NULL\nlast\n";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(csv));
        using var reader = CsvDocument.OpenDataReader(stream, new CsvLoadOptions { NullValue = nullValue });
        using RecordBatch batch = Assert.Single(reader.ReadArrowBatches(new ArrowReadOptions { BatchSize = 10 }));

        Assert.Equal(3, batch.Length);
        Assert.Equal(new[] { "plain", "a\"b\nend", "last" }, ReadStrings(batch, 0));
        Assert.Equal(new[] { string.Empty, string.Empty, string.Empty }, ReadStrings(batch, 1));
        Assert.Equal(new string?[] { "value", nullValue is null ? "NULL" : null, string.Empty }, ReadStrings(batch, 2));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ExplicitTypedCsvBatchesPreserveEveryFieldAndTemporalTick(bool asynchronous) {
        const int rowCount = 257;
        const long firstId = 9007199254740993L;
        DateTime firstDate = new DateTime(2026, 10, 9, 8, 9, 10).AddTicks(1234567);
        string csv = "Name,Id,Date,Value\n" + string.Concat(Enumerable.Range(0, rowCount).Select(row =>
            "Zażółć 😀 " + row.ToString(CultureInfo.InvariantCulture) + "," +
            (firstId + row).ToString(CultureInfo.InvariantCulture) + "," +
            firstDate.AddTicks(row).ToString("O", CultureInfo.InvariantCulture) + "," +
            (row + 0.25d).ToString(CultureInfo.InvariantCulture) + "\n"));
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(csv));
        using var reader = CsvDocument.OpenDataReader(stream);
        List<RecordBatch> batches = await ReadBatches(reader, new ArrowReadOptions {
            BatchSize = 17,
            ColumnTypes = new[] { typeof(string), typeof(long), typeof(DateTime), typeof(double) },
            ColumnNullability = new[] { false, false, false, false },
        }, asynchronous);
        try {
            reader.Close();
            int rowIndex = 0;
            foreach (RecordBatch batch in batches) {
                for (int row = 0; row < batch.Length; row++, rowIndex++) {
                    Assert.Equal("Zażółć 😀 " + rowIndex.ToString(CultureInfo.InvariantCulture),
                        Assert.IsType<StringArray>(batch.Column(0)).GetString(row));
                    Assert.Equal(firstId + rowIndex, Assert.IsType<Int64Array>(batch.Column(1)).GetValue(row));
                    Assert.Equal(firstDate.AddTicks(rowIndex),
                        Assert.IsType<TimestampArray>(batch.Column(2)).GetTimestamp(row)!.Value.DateTime);
                    Assert.Equal(rowIndex + 0.25d, Assert.IsType<DoubleArray>(batch.Column(3)).GetValue(row));
                    for (int column = 0; column < batch.ColumnCount; column++) Assert.False(batch.Column(column).IsNull(row));
                }
            }
            Assert.Equal(rowCount, rowIndex);
        } finally {
            foreach (RecordBatch batch in batches) batch.Dispose();
        }
    }

    private static string?[] ReadStrings(RecordBatch batch, int ordinal) {
        var array = Assert.IsType<StringArray>(batch.Column(ordinal));
        return Enumerable.Range(0, batch.Length).Select(index => array.GetString(index)).ToArray();
    }

    private static async Task<List<RecordBatch>> ReadBatches(DbDataReader reader, ArrowReadOptions options, bool asynchronous) {
        var batches = new List<RecordBatch>();
        try {
            if (asynchronous) {
                await foreach (RecordBatch batch in reader.ReadArrowBatchesAsync(options)) batches.Add(batch);
            } else {
                batches.AddRange(reader.ReadArrowBatches(options));
            }
            return batches;
        } catch {
            foreach (RecordBatch batch in batches) batch.Dispose();
            throw;
        }
    }

    // This provider exposes the public capability without the internal fast-value
    // interface. Its buffer is reused only when advancing or closing the reader.
    private sealed class BorrowedUtf8Reader : DbDataReader, IDataReaderUtf8TextSource {
        private readonly object?[][] _rows;
        private readonly byte[] _buffer = new byte[4096];
        private readonly int[] _starts = new int[5];
        private readonly int[] _lengths = new int[5];
        private readonly Type _lastColumnType;
        private readonly CultureInfo _textCulture;
        private int _row = -1;
        private bool _closed;

        internal BorrowedUtf8Reader(object?[][] rows, Type? lastColumnType = null, CultureInfo? textCulture = null) {
            _rows = rows;
            _lastColumnType = lastColumnType ?? typeof(Uri);
            _textCulture = textCulture ?? CultureInfo.InvariantCulture;
        }
        internal int MaterializedValueCount { get; private set; }
        internal int[] BorrowAttempts { get; } = new int[5];
        public override int FieldCount => 5;
        public override bool HasRows => _rows.Length != 0;
        public override bool IsClosed => _closed;
        public override int Depth => 0;
        public override int RecordsAffected => -1;
        public override object this[int ordinal] => GetValue(ordinal);
        public override object this[string name] => GetValue(GetOrdinal(name));
        public override string GetName(int ordinal) => "Column" + ordinal;
        public override int GetOrdinal(string name) => int.Parse(name.Substring(6), CultureInfo.InvariantCulture);
        public override Type GetFieldType(int ordinal) => ordinal == 4 ? _lastColumnType : typeof(string);
        public override string GetDataTypeName(int ordinal) => GetFieldType(ordinal).Name;
        public override bool NextResult() => false;

        public override bool Read() {
            _buffer.AsSpan().Fill(0xCC);
            if (_closed || ++_row >= _rows.Length) return false;
            int position = 0;
            for (int ordinal = 0; ordinal < FieldCount; ordinal++) {
                _starts[ordinal] = position;
                _lengths[ordinal] = Encoding.UTF8.GetBytes(
                    Convert.ToString(_rows[_row][ordinal], _textCulture).AsSpan(), _buffer.AsSpan(position));
                position += _lengths[ordinal];
            }
            return true;
        }

        public bool TryGetUtf8Text(int ordinal, out ReadOnlySpan<byte> text) {
            BorrowAttempts[ordinal]++;
            if (ordinal == 3) {
                text = default;
                return false;
            }
            text = _buffer.AsSpan(_starts[ordinal], _lengths[ordinal]);
            return true;
        }

        public override void Close() {
            _closed = true;
            _buffer.AsSpan().Fill(0xCC);
        }

        public override bool IsDBNull(int ordinal) => _rows[_row][ordinal] is null;
        public override object GetValue(int ordinal) {
            MaterializedValueCount++;
            return _rows[_row][ordinal] ?? DBNull.Value;
        }
        public override int GetValues(object[] values) {
            int count = Math.Min(values.Length, FieldCount);
            for (int ordinal = 0; ordinal < count; ordinal++) values[ordinal] = GetValue(ordinal);
            return count;
        }
        public override string GetString(int ordinal) => (string)GetValue(ordinal);
        public override bool GetBoolean(int ordinal) => (bool)GetValue(ordinal);
        public override byte GetByte(int ordinal) => (byte)GetValue(ordinal);
        public override char GetChar(int ordinal) => (char)GetValue(ordinal);
        public override DateTime GetDateTime(int ordinal) => (DateTime)GetValue(ordinal);
        public override decimal GetDecimal(int ordinal) => (decimal)GetValue(ordinal);
        public override double GetDouble(int ordinal) => (double)GetValue(ordinal);
        public override float GetFloat(int ordinal) => (float)GetValue(ordinal);
        public override Guid GetGuid(int ordinal) => (Guid)GetValue(ordinal);
        public override short GetInt16(int ordinal) => (short)GetValue(ordinal);
        public override int GetInt32(int ordinal) => (int)GetValue(ordinal);
        public override long GetInt64(int ordinal) => (long)GetValue(ordinal);
        public override IEnumerator GetEnumerator() => throw new NotSupportedException();
        public override long GetBytes(int ordinal, long dataOffset, byte[]? buffer, int bufferOffset, int length) => throw new NotSupportedException();
        public override long GetChars(int ordinal, long dataOffset, char[]? buffer, int bufferOffset, int length) => throw new NotSupportedException();
    }
}
#endif
