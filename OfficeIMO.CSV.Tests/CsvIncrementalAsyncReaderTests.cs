#if NET8_0_OR_GREATER
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvIncrementalAsyncReaderTests
{
    [Theory]
    [InlineData(",", 1)]
    [InlineData("||", 3)]
    [InlineData("||", 4096)]
    public async Task Incremental_Reader_Matches_Canonical_Parsing_Through_Chunk_Boundaries(string delimiter, int chunk)
    {
        string text = $"Name{delimiter}Notes\r\nAlpha{delimiter}\"first\r\nsecond\rthird\nfourth \"\"quoted\"\"\"\r\nBeta{delimiter} last \n";
        var options = new CsvLoadOptions { DelimiterText = delimiter, TrimWhitespace = true, QuoteParsingMode = CsvQuoteParsingMode.Strict };
        using var expectedStream = new MemoryStream(Encoding.UTF8.GetBytes(text));
        using var expected = CsvDocument.OpenDataReader(expectedStream, options);
        using var input = new AsyncInput(Encoding.UTF8.GetBytes(text), chunk);
        using var actual = await CsvDocument.OpenStreamingDataReaderAsync(input, options);
        Assert.Equal(delimiter, ((ICsvDataReaderDialectMetadata)actual).DelimiterText);
        while (expected.Read())
        {
            Assert.True(await actual.ReadAsync());
            for (int i = 0; i < expected.FieldCount; i++) Assert.Equal(expected.GetValue(i), actual.GetValue(i));
        }
        Assert.False(await actual.ReadAsync());
        Assert.True(input.CanRead);
    }

    [Fact]
    public async Task Returns_First_Row_Without_End_Of_Input_And_Cancels_A_Blocked_Advance()
    {
        using var input = new AsyncInput(Encoding.UTF8.GetBytes("Name\nAlpha\n"), 4096, blockAtEnd: true);
        using var reader = await CsvDocument.OpenStreamingDataReaderAsync(input);
        Assert.True(await reader.ReadAsync());
        Assert.Equal("Alpha", reader.GetString(0));
        Assert.Equal(1, ((ICsvDataReaderPositionMetadata)reader).RecordNumber);
        using var cancellation = new CancellationTokenSource();
        Task<bool> advance = reader.ReadAsync(cancellation.Token);
        await input.Blocked.Task.WaitAsync(TimeSpan.FromSeconds(5));
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => advance.WaitAsync(TimeSpan.FromSeconds(5)));
        Assert.Equal(0, ((ICsvDataReaderPositionMetadata)reader).RecordNumber);
        await Assert.ThrowsAsync<InvalidOperationException>(() => reader.ReadAsync());
        reader.Dispose();
        Assert.True(input.CanRead);
    }

    [Fact]
    public async Task Opening_Token_Remains_Active_After_Initialization()
    {
        using var cancellation = new CancellationTokenSource();
        using var input = new AsyncInput(Encoding.UTF8.GetBytes("Name\nAlpha\n"), 4096, blockAtEnd: true);
        using var reader = await CsvDocument.OpenStreamingDataReaderAsync(input, cancellationToken: cancellation.Token);
        Assert.True(await reader.ReadAsync());
        Task<bool> advance = reader.ReadAsync();
        await input.Blocked.Task.WaitAsync(TimeSpan.FromSeconds(5));
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => advance.WaitAsync(TimeSpan.FromSeconds(5)));
    }

    [Fact]
    public async Task Schema_Samples_Are_Bounded_Replayed_And_Keep_Source_Positions()
    {
        using var input = new AsyncInput(Encoding.UTF8.GetBytes("Id,Notes\n1,\"a\r\nb\"\n2,last\n"), 1, blockAtEnd: true);
        using var reader = await CsvDocument.OpenStreamingDataReaderAsync(input,
            readerOptions: new CsvDataReaderOptions { InferSchema = true, SchemaSampleSize = 2 });
        Assert.Equal(typeof(int), reader.GetFieldType(0));
        Assert.True(await reader.ReadAsync());
        Assert.Equal(1, reader.GetInt32(0));
        var metadata = (ICsvDataReaderPositionMetadata)reader;
        Assert.Equal(2, metadata.PhysicalLineNumber);
        Assert.Equal(3, metadata.PhysicalEndLineNumber);
        Assert.True(await reader.ReadAsync());
        Assert.Equal(2, reader.GetInt32(0));
        Assert.Equal(4, metadata.PhysicalLineNumber);
        Assert.False(input.Blocked.Task.IsCompleted);
    }

    [Theory]
    [InlineData(CsvCompressionType.None)]
    [InlineData(CsvCompressionType.GZip)]
    [InlineData(CsvCompressionType.Deflate)]
    public async Task Supports_Bom_Compression_And_Caller_Ownership(CsvCompressionType compression)
    {
        using var bytes = new MemoryStream();
        await CsvDocument.Parse("Name,Value\nAlpha,7\n").SaveAsync(bytes,
            new CsvSaveOptions { Encoding = new UnicodeEncoding(false, true), CompressionType = compression });
        using var input = new AsyncInput(bytes.ToArray(), 2);
        using (var reader = await CsvDocument.OpenStreamingDataReaderAsync(input, new CsvLoadOptions { CompressionType = compression }))
        {
            Assert.Equal("Name", reader.GetName(0));
            Assert.True(await reader.ReadAsync());
            Assert.Equal("7", reader.GetString(1));
            Assert.False(await reader.ReadAsync());
        }
        Assert.True(input.CanRead);
    }

    [Theory]
    [InlineData("# ignored \"\nName,Value\nAlpha,1\n")]
    [InlineData("#Fields: Name Value\nAlpha,1\n")]
    public async Task Handles_Comment_Replay_And_W3c_Headers(string text)
    {
        using var input = new AsyncInput(Encoding.UTF8.GetBytes(text), 1);
        using var reader = await CsvDocument.OpenStreamingDataReaderAsync(input);
        Assert.Equal("Name", reader.GetName(0));
        Assert.Equal("Value", reader.GetName(1));
        Assert.True(await reader.ReadAsync());
        Assert.Equal("Alpha", reader.GetString(0));
        Assert.False(await reader.ReadAsync());
    }

    [Fact]
    public async Task Explicit_Header_Skips_Records_And_Adds_Static_Columns()
    {
        using var input = new AsyncInput(Encoding.UTF8.GetBytes("preamble\n1\n2,last\n"), 2);
        using var reader = await CsvDocument.OpenStreamingDataReaderAsync(input, new CsvLoadOptions
        {
            Header = new[] { "Id", "Notes" }, SkipInitialRecords = 1,
            StaticColumns = new Dictionary<string, object?> { ["Source"] = "fixture" }
        });
        Assert.Equal(3, reader.FieldCount);
        Assert.True(await reader.ReadAsync());
        Assert.Equal("", reader.GetString(1));
        Assert.Equal("fixture", reader.GetString(2));
        Assert.True(await reader.ReadAsync());
        Assert.Equal("last", reader.GetString(1));
        Assert.False(await reader.ReadAsync());
    }

    [Fact]
    public async Task Strict_Parse_Errors_And_Input_Limits_Do_Not_Close_Caller_Stream()
    {
        using var input = new AsyncInput(Encoding.UTF8.GetBytes("Name\n\"unfinished"), 1);
        using (var reader = await CsvDocument.OpenStreamingDataReaderAsync(input, new CsvLoadOptions { QuoteParsingMode = CsvQuoteParsingMode.Strict }))
            await Assert.ThrowsAsync<CsvParseException>(() => reader.ReadAsync());
        Assert.True(input.CanRead);
        using var limited = new AsyncInput(Encoding.UTF8.GetBytes("Name\nAlpha\n"), 1);
        using (var reader = await CsvDocument.OpenStreamingDataReaderAsync(limited, new CsvLoadOptions { MaxInputBytes = 8 }))
            await Assert.ThrowsAsync<InvalidDataException>(() => reader.ReadAsync());
        Assert.True(limited.CanRead);
    }

    [Fact]
    public async Task Can_Mix_Synchronous_Lookahead_With_Async_And_Sync_Advances()
    {
        using var input = new MemoryStream(Encoding.UTF8.GetBytes("Name\nAlpha\nBeta\nGamma\n"));
        using var reader = await CsvDocument.OpenStreamingDataReaderAsync(input);
        Assert.True(reader.HasRows);
        Assert.True(await reader.ReadAsync());
        Assert.Equal("Alpha", reader.GetString(0));
        Assert.True(reader.Read());
        Assert.Equal("Beta", reader.GetString(0));
        Assert.True(await reader.ReadAsync());
        Assert.Equal("Gamma", reader.GetString(0));
        Assert.Equal(3, ((ICsvDataReaderPositionMetadata)reader).RecordNumber);
        Assert.False(await reader.ReadAsync());
    }

    [Theory]
    [InlineData("")]
    [InlineData("1,Alpha\n2,Beta\n")]
    public async Task Headerless_Input_Replays_Its_First_Record(string text)
    {
        using var input = new AsyncInput(Encoding.UTF8.GetBytes(text), 1);
        using var reader = await CsvDocument.OpenStreamingDataReaderAsync(input, new CsvLoadOptions { HasHeaderRow = false });
        if (text.Length == 0) { Assert.Equal(0, reader.FieldCount); Assert.False(await reader.ReadAsync()); return; }
        Assert.Equal(2, reader.FieldCount);
        Assert.True(await reader.ReadAsync());
        Assert.Equal("1", reader.GetString(0));
        Assert.True(await reader.ReadAsync());
        Assert.Equal("2", reader.GetString(0));
        Assert.False(await reader.ReadAsync());
    }

    [Fact]
    public async Task Static_Null_And_Typed_Values_Match_Existing_Reader_Projection()
    {
        var options = new CsvLoadOptions
        {
            StaticColumns = new Dictionary<string, object?> { ["Missing"] = null, ["Count"] = 7, ["Active"] = true }
        };
        using var expectedInput = new MemoryStream(Encoding.UTF8.GetBytes("Name\nAlpha\n"));
        using var expected = CsvDocument.OpenDataReader(expectedInput, options);
        using var input = new AsyncInput(Encoding.UTF8.GetBytes("Name\nAlpha\n"), 1);
        using var actual = await CsvDocument.OpenStreamingDataReaderAsync(input, options);
        Assert.True(expected.Read());
        Assert.True(await actual.ReadAsync());
        for (int i = 0; i < expected.FieldCount; i++)
        {
            Assert.Equal(expected.GetValue(i), actual.GetValue(i));
            Assert.Equal(expected.IsDBNull(i), actual.IsDBNull(i));
        }
        var expectedValues = new object[expected.FieldCount];
        var actualValues = new object[actual.FieldCount];
        expected.GetValues(expectedValues);
        actual.GetValues(actualValues);
        Assert.Equal(expectedValues, actualValues);
    }

    [Fact]
    public async Task Explicit_Fallback_Restrictions_Are_Checked_Before_Reading()
    {
        using var input = new AsyncInput(Array.Empty<byte>(), 1, blockAtEnd: true);
        await Assert.ThrowsAsync<NotSupportedException>(() => CsvDocument.OpenStreamingDataReaderAsync(input,
            new CsvLoadOptions { DetectDelimiter = true }));
        await Assert.ThrowsAsync<NotSupportedException>(() => CsvDocument.OpenStreamingDataReaderAsync(input,
            readerOptions: new CsvDataReaderOptions { ParallelProcessing = new CsvDataReaderParallelOptions() }));
        Assert.False(input.Blocked.Task.IsCompleted);
        Assert.True(input.CanRead);
    }

    [Fact]
    public async Task Compressed_Path_Enforces_Compressed_Input_Limit_And_Releases_The_File()
    {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-csv-incremental-" + Guid.NewGuid().ToString("N") + ".csv.gz");
        try
        {
            await CsvDocument.Parse("Name\nAlpha\n").SaveAsync(path, new CsvSaveOptions { CompressionType = CsvCompressionType.GZip });
            await Assert.ThrowsAsync<InvalidDataException>(() => CsvDocument.OpenStreamingDataReaderAsync(path, new CsvLoadOptions { MaxInputBytes = 2 }));
            using var exclusive = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        }
        finally { File.Delete(path); }
    }

    private sealed class AsyncInput : Stream
    {
        private readonly byte[] _bytes;
        private readonly int _chunk;
        private readonly bool _block;
        private int _position;
        private bool _disposed;
        internal TaskCompletionSource Blocked { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal AsyncInput(byte[] bytes, int chunk, bool blockAtEnd = false) { _bytes = bytes; _chunk = chunk; _block = blockAtEnd; }
        public override bool CanRead => !_disposed;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) => throw new InvalidOperationException("Synchronous input is forbidden.");
        public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            if (_position == _bytes.Length && _block)
            {
                Blocked.TrySetResult();
                await Task.Delay(Timeout.InfiniteTimeSpan, token);
            }
            int read = Math.Min(Math.Min(count, _chunk), _bytes.Length - _position);
            Array.Copy(_bytes, _position, buffer, offset, read);
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
