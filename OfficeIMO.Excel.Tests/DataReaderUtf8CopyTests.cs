#if NET8_0_OR_GREATER
using System.Data;
using System.Text;
using OfficeIMO.Data;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public sealed class DataReaderUtf8CopyTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Utf8Copy_QueriesAndCopiesByteSegmentsIncludingPartialCharacters(bool borrowed) {
        const string text = "AŻ🐢終Z";
        byte[] expected = Encoding.UTF8.GetBytes(text);
        IDataRecord record = borrowed ? new BorrowedTextRecord(expected) : CreateStringReader(text);
        try {
            Assert.Equal(expected.Length, record.GetUtf8Bytes(0, long.MaxValue, null, 0, 0));
            byte[] destination = Enumerable.Repeat((byte)0xCC, 10).ToArray();
            Assert.Equal(5, record.GetUtf8Bytes(0, 2, destination, 3, 5));
            Assert.Equal(expected.Skip(2).Take(5), destination.Skip(3).Take(5));
            Assert.All(destination.Take(3).Concat(destination.Skip(8)), b => Assert.Equal(0xCC, b));
            Assert.Equal(1, record.GetUtf8Bytes(0, expected.Length - 1, destination, 0, 5));
            Assert.Equal(expected[^1], destination[0]);
            Assert.Equal(0, record.GetUtf8Bytes(0, expected.Length, destination, 0, 1));
            Assert.Equal(0, record.GetUtf8Bytes(0, long.MaxValue, destination, 0, 1));
            Assert.Equal(0, record.GetUtf8Bytes(0, 0, destination, destination.Length, 0));
        } finally {
            (record as IDisposable)?.Dispose();
        }
    }

    [Fact]
    public void Utf8Copy_FallbackPreservesEncodingAcrossChunksAndInvalidSurrogates() {
        string text = new string('a', 4095) + "🐢" + new string('終', 3000) + "\uD800Z\uDC00";
        byte[] expected = Encoding.UTF8.GetBytes(text);
        using DataTableReader reader = CreateStringReader(text);
        Assert.Equal(expected.Length, reader.GetUtf8Bytes(0, 0, null, 0, 0));
        byte[] all = new byte[expected.Length];
        Assert.Equal(expected.Length, reader.GetUtf8Bytes(0, 0, all, 0, all.Length));
        Assert.Equal(expected, all);
        byte[] segment = new byte[13];
        Assert.Equal(segment.Length, reader.GetUtf8Bytes(0, 4096, segment, 0, segment.Length));
        Assert.Equal(expected.Skip(4096).Take(segment.Length), segment);
    }

    [Fact]
    public void Utf8Copy_ValidatesDestinationRangesWithoutOverflow() {
        using DataTableReader reader = CreateStringReader("text");
        byte[] destination = { 0xA1, 0xA2, 0xA3 };
        Assert.Throws<ArgumentNullException>(() => DataReaderUtf8TextExtensions.GetUtf8Bytes(null!, 0, 0, null, 0, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => reader.GetUtf8Bytes(0, -1, null, 0, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => reader.GetUtf8Bytes(0, 0, null, -1, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => reader.GetUtf8Bytes(0, 0, null, 0, -1));
        Assert.Throws<ArgumentOutOfRangeException>(() => reader.GetUtf8Bytes(0, 0, destination, 4, 0));
        Assert.Throws<ArgumentException>(() => reader.GetUtf8Bytes(0, 0, destination, 1, int.MaxValue));
        Assert.Throws<ArgumentOutOfRangeException>(() => reader.GetUtf8Bytes(0, 0, destination, int.MaxValue, int.MaxValue));
        Assert.Equal(new byte[] { 0xA1, 0xA2, 0xA3 }, destination);
    }

    [Fact]
    public void Utf8Copy_UsesTheProvidersCursorOrdinalAndNullContracts() {
        var table = new DataTable();
        table.Columns.Add("Text", typeof(string));
        table.Rows.Add(string.Empty);
        table.Rows.Add(DBNull.Value);
        using (DataTableReader beforeRead = table.CreateDataReader()) {
            Assert.Throws<InvalidOperationException>(() => beforeRead.GetUtf8Bytes(0, 0, null, 0, 0));
        }
        using DataTableReader reader = table.CreateDataReader();
        Assert.True(reader.Read());
        Assert.Equal(0, reader.GetUtf8Bytes(0, 0, null, 0, 0));
        Assert.Throws<IndexOutOfRangeException>(() => reader.GetUtf8Bytes(-1, 0, null, 0, 0));
        Assert.Throws<IndexOutOfRangeException>(() => reader.GetUtf8Bytes(1, 0, null, 0, 0));
        Assert.True(reader.Read());
        Assert.Throws<InvalidCastException>(() => reader.GetUtf8Bytes(0, 0, null, 0, 0));
        Assert.False(reader.Read());
        Assert.Equal(
            Record.Exception(() => reader.GetString(0))?.GetType(),
            Record.Exception(() => reader.GetUtf8Bytes(0, 0, null, 0, 0))?.GetType());
        reader.Close();
        Assert.Throws<InvalidOperationException>(() => reader.GetUtf8Bytes(0, 0, null, 0, 0));
    }

    [Fact]
    public void Utf8Copy_LeavesBinaryGetBytesContractAvailable() {
        byte[] expected = { 0, 127, 128, 255 };
        var table = new DataTable();
        table.Columns.Add("Binary", typeof(byte[]));
        table.Rows.Add(new object[] { expected });
        using DataTableReader reader = table.CreateDataReader();
        Assert.True(reader.Read());
        Assert.Equal(expected.Length, reader.GetBytes(0, 0, null, 0, 0));
        Assert.Throws<InvalidCastException>(() => reader.GetUtf8Bytes(0, 0, null, 0, 0));
        byte[] destination = new byte[2];
        Assert.Equal(2, reader.GetBytes(0, 1, destination, 0, destination.Length));
        Assert.Equal(new byte[] { 127, 128 }, destination);
    }

    private static DataTableReader CreateStringReader(string text) {
        var table = new DataTable();
        table.Columns.Add("Text", typeof(string));
        table.Rows.Add(text);
        DataTableReader reader = table.CreateDataReader();
        Assert.True(reader.Read());
        return reader;
    }

    // A provider boundary that deliberately offers only borrowed text. Calling any
    // materializing getter is a regression in the extension's borrowed path.
    private sealed class BorrowedTextRecord(byte[] text) : IDataRecord, IDataReaderUtf8TextSource {
        public int FieldCount => 1;
        public bool TryGetUtf8Text(int ordinal, out ReadOnlySpan<byte> value) {
            if (ordinal != 0) throw new IndexOutOfRangeException();
            value = text;
            return true;
        }
        public string GetString(int ordinal) => throw new NotSupportedException();
        public object this[int ordinal] => throw new NotSupportedException();
        public object this[string name] => throw new NotSupportedException();
        public bool GetBoolean(int ordinal) => throw new NotSupportedException();
        public byte GetByte(int ordinal) => throw new NotSupportedException();
        public long GetBytes(int ordinal, long offset, byte[]? buffer, int bufferOffset, int length) => throw new NotSupportedException();
        public char GetChar(int ordinal) => throw new NotSupportedException();
        public long GetChars(int ordinal, long offset, char[]? buffer, int bufferOffset, int length) => throw new NotSupportedException();
        public IDataReader GetData(int ordinal) => throw new NotSupportedException();
        public string GetDataTypeName(int ordinal) => throw new NotSupportedException();
        public DateTime GetDateTime(int ordinal) => throw new NotSupportedException();
        public decimal GetDecimal(int ordinal) => throw new NotSupportedException();
        public double GetDouble(int ordinal) => throw new NotSupportedException();
        public Type GetFieldType(int ordinal) => throw new NotSupportedException();
        public float GetFloat(int ordinal) => throw new NotSupportedException();
        public Guid GetGuid(int ordinal) => throw new NotSupportedException();
        public short GetInt16(int ordinal) => throw new NotSupportedException();
        public int GetInt32(int ordinal) => throw new NotSupportedException();
        public long GetInt64(int ordinal) => throw new NotSupportedException();
        public string GetName(int ordinal) => throw new NotSupportedException();
        public int GetOrdinal(string name) => throw new NotSupportedException();
        public object GetValue(int ordinal) => throw new NotSupportedException();
        public int GetValues(object[] values) => throw new NotSupportedException();
        public bool IsDBNull(int ordinal) => throw new NotSupportedException();
    }
}
#endif
