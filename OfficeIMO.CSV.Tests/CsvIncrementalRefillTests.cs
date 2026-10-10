#if NET8_0_OR_GREATER
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.CSV;
using Xunit;
using AsyncInput = OfficeIMO.CSV.Tests.CsvIncrementalAsyncReaderTests.AsyncInput;

namespace OfficeIMO.CSV.Tests;

public class CsvIncrementalRefillTests
{
    [Theory]
    [InlineData(3, true)]
    [InlineData(32768, true)]
    [InlineData(32768, false)]
    public async Task RefillsPreserveFieldsPositionsAndRetainedStringsAcrossFallbacks(int chunk, bool asynchronous)
    {
        const int boundary = 32768;
        var csv = new StringBuilder("Id,Text,Empty,Nullable,Missing\n");
        string padding = new string('p', boundary - 8 - csv.Length - "0,,,NULL,\n".Length);
        csv.Append("0,").Append(padding).Append(",,NULL,\n");
        // This short row begins eight ASCII bytes before a full transport read ends.
        csv.Append("1,  crossing  ,,NULL,\n2,");
        // The next quoted field begins only after another deliberately short read ends.
        int quoteBoundary = csv.Length;
        csv.Append("\"one\r\ntwo \"\"quoted\"\"\",,,\r\n");
        string oversized = new string('x', 40000);
        csv.Append("3,").Append(oversized).Append(",,NULL,\n# ignored,,NULL,\n4, final ,,NULL");
        byte[] bytes = Encoding.UTF8.GetBytes(csv.ToString());
        var options = new CsvLoadOptions
        {
            TrimWhitespace = true, NullValue = "NULL", SkipCommentRows = true,
            QuoteParsingMode = CsvQuoteParsingMode.Strict
        };
        using var expectedInput = new MemoryStream(bytes);
        using var expected = CsvDocument.OpenDataReader(expectedInput, options);
        using var input = new SegmentedInput(bytes, chunk, !asynchronous, boundary, quoteBoundary, bytes.Length);
        using var actual = (CsvDataReader)await CsvDocument.OpenDataReaderAsync(input, options);
        string[] names = { padding, "crossing", "one\r\ntwo \"quoted\"", oversized, "final" };
        int[] startLines = { 2, 3, 4, 6, 8 };
        int[] endLines = { 2, 3, 5, 6, 8 };
        var retained = new List<string>();
        int row = 0;
        while (expected.Read())
        {
            Assert.True(asynchronous ? await actual.ReadAsync() : actual.Read());
            Assert.Equal(row, actual.GetInt32(0));
            for (int column = 0; column < expected.FieldCount; column++)
            {
                Assert.Equal(expected.GetValue(column), actual.GetValue(column));
                Assert.Equal(expected.IsDBNull(column), actual.IsDBNull(column));
            }
            var position = (ICsvDataReaderPositionMetadata)actual;
            Assert.Equal(startLines[row], position.PhysicalLineNumber);
            Assert.Equal(endLines[row], position.PhysicalEndLineNumber);
            Assert.Equal(row + 1, position.RecordNumber);
            AssertFieldStates(actual, row);
            retained.Add(actual.GetString(1));
            row++;
        }
        Assert.Equal(names.Length, row);
        Assert.False(asynchronous ? await actual.ReadAsync() : actual.Read());
        actual.Dispose();
        Assert.Equal(names, retained);
        Assert.True(input.AsyncReads > 0);
        if (asynchronous) Assert.Equal(0, input.SynchronousReads);
        else Assert.True(input.SynchronousReads > 0);
        Assert.True(input.CanRead);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ShortRowEndingAtBoundaryCrPreservesItsFollowingRecord(bool lineFeed)
    {
        const string row = "crossing\r";
        var csv = new StringBuilder("V\n");
        string padding = new string('p', 32768 - row.Length - csv.Length - 1);
        csv.Append(padding).Append('\n').Append(row);
        if (lineFeed) csv.Append('\n');
        csv.Append("# ignored\r\nnext\r\nlast");
        using var input = new AsyncInput(Encoding.UTF8.GetBytes(csv.ToString()), 32768);
        using var reader = await CsvDocument.OpenDataReaderAsync(input, new CsvLoadOptions { SkipCommentRows = true });
        Assert.True(await reader.ReadAsync());
        Assert.Equal(padding, reader.GetString(0));
        Assert.True(await reader.ReadAsync());
        string retained = reader.GetString(0);
        Assert.Equal("crossing", retained);
        Assert.Equal(3, ((ICsvDataReaderPositionMetadata)reader).PhysicalLineNumber);
        Assert.True(await reader.ReadAsync());
        Assert.Equal("next", reader.GetString(0));
        Assert.Equal(5, ((ICsvDataReaderPositionMetadata)reader).PhysicalLineNumber);
        Assert.True(await reader.ReadAsync());
        Assert.Equal("last", reader.GetString(0));
        Assert.Equal(6, ((ICsvDataReaderPositionMetadata)reader).PhysicalLineNumber);
        Assert.False(await reader.ReadAsync());
        Assert.Equal("crossing", retained);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CancellationDuringPartialTailRefillHidesPreviousRowAndEndsReader(bool lifetimeCancellation)
    {
        string csv = "V\nfirst\n" + new string(' ', 32760 - "V\nfirst\n".Length - 1) + "\npartial!";
        using var input = new AsyncInput(Encoding.UTF8.GetBytes(csv), 32768, blockAtEnd: true);
        using var lifetime = new CancellationTokenSource();
        using var operation = new CancellationTokenSource();
        using var reader = await CsvDocument.OpenDataReaderAsync(input,
            new CsvLoadOptions { TrimWhitespace = true }, cancellationToken: lifetime.Token);
        Assert.True(await reader.ReadAsync());
        string retained = reader.GetString(0);
        Assert.Equal("first", retained);
        Task<bool> read = reader.ReadAsync(operation.Token);
        await input.Blocked.Task.WaitAsync(TimeSpan.FromSeconds(5));
        if (lifetimeCancellation) lifetime.Cancel();
        else operation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => read.WaitAsync(TimeSpan.FromSeconds(5)));
        Assert.Equal(0, ((ICsvDataReaderPositionMetadata)reader).RecordNumber);
        Assert.Throws<InvalidOperationException>(() => reader.GetString(0));
        await Assert.ThrowsAsync<InvalidOperationException>(() => reader.ReadAsync());
        Assert.Equal("first", retained);
        Assert.True(input.CanRead);
    }

    private static void AssertFieldStates(CsvDataReader reader, int id)
    {
        var record = new CsvRecord(reader);
        Assert.True(record.GetSpan(2).IsEmpty);
        Assert.False(record.IsMissing(2));
        Assert.False(record.IsNull(2));
        Assert.Equal(id != 2, record.IsNull(3));
        Assert.False(record.IsMissing(3));
        Assert.Equal(id == 4, record.IsMissing(4));
        Assert.False(record.IsNull(4));
        Assert.True(record.GetSpan(4).IsEmpty);
    }

    // A stream boundary fake makes the two split locations deterministic while exercising
    // the public reader, BOM/byte-limit wrappers, decoder and canonical field parser.
    private sealed class SegmentedInput : Stream
    {
        private readonly byte[] _bytes;
        private readonly int[] _ends;
        private readonly int _chunk;
        private readonly bool _allowSynchronous;
        private int _position, _segment;
        private bool _disposed;
        internal int AsyncReads { get; private set; }
        internal int SynchronousReads { get; private set; }
        internal SegmentedInput(byte[] bytes, int chunk, bool allowSynchronous, params int[] ends)
        {
            _bytes = bytes; _chunk = chunk; _allowSynchronous = allowSynchronous; _ends = ends;
        }
        public override bool CanRead => !_disposed;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count)
        {
            if (!_allowSynchronous) throw new InvalidOperationException("Synchronous input is forbidden.");
            SynchronousReads++;
            return ReadChunk(buffer.AsSpan(offset, count), CancellationToken.None);
        }
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken token)
        {
            AsyncReads++;
            return Task.FromResult(ReadChunk(buffer.AsSpan(offset, count), token));
        }
        public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken token = default)
        {
            AsyncReads++;
            return new ValueTask<int>(ReadChunk(buffer.Span, token));
        }
        private int ReadChunk(Span<byte> buffer, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            if (_position == _bytes.Length || buffer.IsEmpty) return 0;
            while (_position == _ends[_segment]) _segment++;
            int read = Math.Min(Math.Min(buffer.Length, _chunk), _ends[_segment] - _position);
            _bytes.AsSpan(_position, read).CopyTo(buffer);
            _position += read;
            return read;
        }
        public override void Flush() => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) { _disposed = true; base.Dispose(disposing); }
    }
}
#endif
