#if NET8_0_OR_GREATER
using System.Data.Common;
using System.Linq;
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
    [InlineData(",", 4096)]
    [InlineData("||", 3)]
    [InlineData("||", 4096)]
    public async Task Incremental_Reader_Matches_Canonical_Parsing_Through_Chunk_Boundaries(string delimiter, int chunk)
    {
        string text = $"Name{delimiter}Notes\r\nAlpha{delimiter}\"first\r\nsecond\rthird\nfourth \"\"quoted\"\"\"\r\nBeta{delimiter} last \n";
        var options = new CsvLoadOptions { DelimiterText = delimiter, TrimWhitespace = true, QuoteParsingMode = CsvQuoteParsingMode.Strict };
        using var expectedStream = new MemoryStream(Encoding.UTF8.GetBytes(text));
        using var expected = CsvDocument.OpenDataReader(expectedStream, options);
        using var input = new AsyncInput(Encoding.UTF8.GetBytes(text), chunk);
        using var actual = await CsvDocument.OpenDataReaderAsync(input, options);
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
        using var reader = await CsvDocument.OpenDataReaderAsync(input);
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
        using var reader = await CsvDocument.OpenDataReaderAsync(input, cancellationToken: cancellation.Token);
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
        using var reader = await CsvDocument.OpenDataReaderAsync(input,
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
        using (var reader = await CsvDocument.OpenDataReaderAsync(input, new CsvLoadOptions { CompressionType = compression }))
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
        using var reader = await CsvDocument.OpenDataReaderAsync(input);
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
        using var reader = await CsvDocument.OpenDataReaderAsync(input, new CsvLoadOptions
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
        using (var reader = await CsvDocument.OpenDataReaderAsync(input, new CsvLoadOptions { QuoteParsingMode = CsvQuoteParsingMode.Strict }))
            await Assert.ThrowsAsync<CsvParseException>(() => reader.ReadAsync());
        Assert.True(input.CanRead);
        using var limited = new AsyncInput(Encoding.UTF8.GetBytes("Name\nAlpha\n"), 1);
        using (var reader = await CsvDocument.OpenDataReaderAsync(limited, new CsvLoadOptions { MaxInputBytes = 8 }))
            await Assert.ThrowsAsync<InvalidDataException>(() => reader.ReadAsync());
        Assert.True(limited.CanRead);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Can_Mix_Lookahead_And_Advances_Without_Changing_Retained_Values(bool inferSchema)
    {
        using var input = new MemoryStream(Encoding.UTF8.GetBytes("Name\nAlpha\nBeta\n\"Gamma \"\"quoted\"\"\"\nDelta\n"));
        using var reader = await CsvDocument.OpenDataReaderAsync(input,
            readerOptions: new CsvDataReaderOptions { InferSchema = inferSchema, SchemaSampleSize = 2 });
        Assert.True(reader.HasRows);
        Assert.True(await reader.ReadAsync());
        Assert.Equal("Alpha", reader.GetString(0));
        var retainedSample = new object[1];
        Assert.Equal(1, reader.GetValues(retainedSample));
        Assert.True(reader.Read());
        Assert.Equal("Beta", reader.GetString(0));
        Assert.True(await reader.ReadAsync());
        Assert.Equal("Gamma \"quoted\"", reader.GetString(0));
        Assert.Equal(3, ((ICsvDataReaderPositionMetadata)reader).RecordNumber);
        var retainedLiveRow = new object[1];
        Assert.Equal(1, reader.GetValues(retainedLiveRow));
        Assert.True(reader.Read());
        Assert.Equal("Delta", reader.GetString(0));
        Assert.Equal("Alpha", retainedSample[0]);
        Assert.Equal("Gamma \"quoted\"", retainedLiveRow[0]);
        Assert.False(await reader.ReadAsync());
    }

    [Theory]
    [InlineData("")]
    [InlineData("1,Alpha\n2,Beta\n")]
    public async Task Headerless_Input_Replays_Its_First_Record(string text)
    {
        using var input = new AsyncInput(Encoding.UTF8.GetBytes(text), 1);
        using var reader = await CsvDocument.OpenDataReaderAsync(input, new CsvLoadOptions { HasHeaderRow = false });
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
        using var actual = await CsvDocument.OpenDataReaderAsync(input, options);
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
    public async Task Invalid_Parallel_Options_Are_Checked_Before_Reading()
    {
        using var input = new AsyncInput(Array.Empty<byte>(), 1, blockAtEnd: true);
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => CsvDocument.OpenDataReaderAsync(input,
            readerOptions: new CsvDataReaderOptions { ParallelProcessing = new CsvDataReaderParallelOptions { BatchSize = 0 } }));
        Assert.False(input.Blocked.Task.IsCompleted);
        Assert.True(input.CanRead);
    }

    [Fact]
    public async Task Parallel_Projection_Uses_Async_Input_And_Preserves_Order_And_Positions()
    {
        var csv = new StringBuilder("Id;Name\r\n");
        for (int id = 1; id <= 1030; id++) csv.Append(id).Append(";\"row ").Append(id).Append("\r\nend\"\r\n");
        using var input = new AsyncInput(Encoding.UTF8.GetBytes(csv.ToString()), 3);
        var schema = new CsvSchemaBuilder().Column("Id").AsInt32().Column("Name").AsString().Done().Build();
        using var reader = await CsvDocument.OpenDataReaderAsync(input,
            new CsvLoadOptions { DetectDelimiter = true }, new CsvDataReaderOptions
            {
                Schema = schema,
                ParallelProcessing = new CsvDataReaderParallelOptions { BatchSize = 127, MaxDegreeOfParallelism = 3 }
            });
        int expected = 0;
        var position = (ICsvDataReaderPositionMetadata)reader;
        while (await reader.ReadAsync())
        {
            expected++;
            Assert.Equal(expected, reader.GetInt32(0));
            Assert.Equal($"row {expected}\r\nend", reader.GetString(1));
            Assert.Equal(expected, position.RecordNumber);
            Assert.Equal(expected * 2, position.PhysicalLineNumber);
            Assert.Equal(expected * 2 + 1, position.PhysicalEndLineNumber);
            Assert.True(reader.HasRows);
        }
        Assert.Equal(1030, expected);
        Assert.Equal(0, position.RecordNumber);
        Assert.True(input.CanRead);
    }

    [Fact]
    public async Task Parallel_Cancellation_Interrupts_Blocked_Async_Batch_Capture()
    {
        using var input = new AsyncInput(Encoding.UTF8.GetBytes("Id\n1\n"), 1, blockAtEnd: true);
        using var reader = await CsvDocument.OpenDataReaderAsync(input, readerOptions: new CsvDataReaderOptions
        {
            Schema = new CsvSchemaBuilder().Column("Id").AsInt32().Done().Build(),
            ParallelProcessing = new CsvDataReaderParallelOptions { BatchSize = 2, MaxDegreeOfParallelism = 2 }
        });
        using var cancellation = new CancellationTokenSource();
        Task<bool> reading = reader.ReadAsync(cancellation.Token);
        await input.Blocked.Task.WaitAsync(TimeSpan.FromSeconds(10));
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => reading);
        Assert.Equal(0, ((ICsvDataReaderPositionMetadata)reader).RecordNumber);
        Assert.Throws<InvalidOperationException>(() => reader.GetValue(0));
        await Assert.ThrowsAsync<InvalidOperationException>(() => reader.ReadAsync());
        Assert.True(input.CanRead);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Parallel_Projection_Error_Yields_Valid_Prefix_Then_Ends_Reader(bool synchronous)
    {
        using var input = new MemoryStream(Encoding.UTF8.GetBytes("Id\n1\n2\nbad\n4\n"));
        using var reader = await CsvDocument.OpenDataReaderAsync(input, readerOptions: new CsvDataReaderOptions
        {
            Schema = new CsvSchemaBuilder().Column("Id").AsInt32().Done().Build(),
            ParallelProcessing = new CsvDataReaderParallelOptions { BatchSize = 2, MaxDegreeOfParallelism = 3 }
        });
        Assert.True(await reader.ReadAsync());
        Assert.Equal(1, reader.GetInt32(0));
        Assert.True(await reader.ReadAsync());
        Assert.Equal(2, reader.GetInt32(0));
        if (synchronous) Assert.ThrowsAny<CsvException>(() => reader.Read());
        else await Assert.ThrowsAnyAsync<CsvException>(() => reader.ReadAsync());
        Assert.Equal(0, ((ICsvDataReaderPositionMetadata)reader).RecordNumber);
        Assert.Throws<InvalidOperationException>(() => reader.GetValue(0));
        await Assert.ThrowsAsync<InvalidOperationException>(() => reader.ReadAsync());
        Assert.Throws<InvalidOperationException>(() => reader.Read());
    }

    [Theory]
    [InlineData(1)]
    [InlineData(3)]
    [InlineData(4096)]
    public async Task Detection_Replays_Comments_Skips_And_Multiline_Records(int chunk)
    {
        const string csv = "# generated \"by tool\r\nignored\r\nName;Value\r\nAlpha;\"one\r\ntwo,three\"\r\nBeta;7\r\n";
        var options = new CsvLoadOptions { DetectDelimiter = true, SkipInitialRecords = 1 };
        using var expectedInput = new MemoryStream(Encoding.UTF8.GetBytes(csv));
        using var expected = CsvDocument.OpenDataReader(expectedInput, options);
        using var input = new AsyncInput(Encoding.UTF8.GetBytes(csv), chunk);
        using var actual = await CsvDocument.OpenDataReaderAsync(input, options);
        Assert.Equal(';', ((ICsvDataReaderMetadata)actual).Delimiter);
        Assert.True(options.DetectDelimiter);
        Assert.Equal(',', options.Delimiter);
        int row = 0;
        while (expected.Read())
        {
            Assert.True(await actual.ReadAsync());
            for (int i = 0; i < expected.FieldCount; i++) Assert.Equal(expected.GetValue(i), actual.GetValue(i));
            var position = (ICsvDataReaderPositionMetadata)actual;
            Assert.Equal(row++ == 0 ? 4 : 6, position.PhysicalLineNumber);
        }
        Assert.False(await actual.ReadAsync());
    }

    [Fact]
    public async Task Detection_Returns_Before_Unbounded_Source_Eof()
    {
        string csv = "Name;Value\n" + string.Concat(Enumerable.Repeat("Alpha;7\n", 20_000));
        using var input = new AsyncInput(Encoding.UTF8.GetBytes(csv), 4096, blockAtEnd: true);
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(10));
        using var reader = await CsvDocument.OpenDataReaderAsync(input,
            new CsvLoadOptions { DetectDelimiter = true }, cancellationToken: timeout.Token);
        Assert.True(await reader.ReadAsync(timeout.Token));
        Assert.Equal("Alpha", reader.GetString(0));
        Assert.Equal("7", reader.GetString(1));
        Assert.False(input.Blocked.Task.IsCompleted);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(4096)]
    public async Task Detection_Completes_With_Enough_Records_After_Unmatched_Quote_Comment(int chunk)
    {
        string csv = "# generated \"by tool\nName;Value\n" + string.Concat(Enumerable.Repeat("Alpha;7\n", 100));
        using var input = new AsyncInput(Encoding.UTF8.GetBytes(csv), chunk, blockAtEnd: true);
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(2));
        using var reader = await CsvDocument.OpenDataReaderAsync(input,
            new CsvLoadOptions { DetectDelimiter = true }, cancellationToken: timeout.Token);
        Assert.True(await reader.ReadAsync(timeout.Token));
        Assert.Equal("Alpha", reader.GetString(0));
        Assert.Equal("7", reader.GetString(1));
        Assert.False(input.Blocked.Task.IsCompleted);
    }

    [Fact]
    public async Task Detection_Completes_Last_Sample_After_Many_Quoted_Continuation_Lines()
    {
        string multiline = string.Concat(Enumerable.Repeat("line\n", 1000)) + "end";
        string csv = "Name;Value\n" + string.Concat(Enumerable.Repeat("Alpha;7\n", 62)) + "Beta;\"" + multiline + "\"\n";
        using var input = new AsyncInput(Encoding.UTF8.GetBytes(csv), 1, blockAtEnd: true);
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(10));
        using var reader = await CsvDocument.OpenDataReaderAsync(input,
            new CsvLoadOptions { DetectDelimiter = true }, cancellationToken: timeout.Token);
        for (int row = 0; row < 62; row++)
        {
            Assert.True(await reader.ReadAsync(timeout.Token));
            Assert.Equal("Alpha", reader.GetString(0));
        }
        Assert.True(await reader.ReadAsync(timeout.Token));
        Assert.Equal("Beta", reader.GetString(0));
        Assert.Equal(multiline, reader.GetString(1));
        Assert.False(input.Blocked.Task.IsCompleted);
    }

    [Fact]
    public async Task Detection_Cap_Inside_Record_Uses_Fallback_And_Replays_Whole_Field()
    {
        string name = new string('x', 70_000);
        using var input = new AsyncInput(Encoding.UTF8.GetBytes(name + "|7\n"), 4096);
        using var reader = await CsvDocument.OpenDataReaderAsync(input,
            new CsvLoadOptions { DetectDelimiter = true, Delimiter = '|', Header = new[] { "Name", "Value" } });
        Assert.True(await reader.ReadAsync());
        Assert.Equal(name, reader.GetString(0));
        Assert.Equal("7", reader.GetString(1));
        Assert.False(await reader.ReadAsync());
    }

    [Fact]
    public async Task Detection_Cancellation_Leaves_Caller_Stream_Open()
    {
        using var input = new AsyncInput(Array.Empty<byte>(), 1, blockAtEnd: true);
        using var cancellation = new CancellationTokenSource();
        Task<DbDataReader> opening = CsvDocument.OpenDataReaderAsync(input,
            new CsvLoadOptions { DetectDelimiter = true }, cancellationToken: cancellation.Token);
        await input.Blocked.Task.WaitAsync(TimeSpan.FromSeconds(10));
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => opening);
        Assert.True(input.CanRead);
    }

    [Fact]
    public async Task Explicit_Text_Delimiter_Remains_Effective_When_Detection_Is_Enabled()
    {
        using var input = new AsyncInput(Encoding.UTF8.GetBytes("Name||Value\nAlpha||7\n"), 1);
        using var reader = await CsvDocument.OpenDataReaderAsync(input,
            new CsvLoadOptions { DetectDelimiter = true, DelimiterText = "||" });
        Assert.Equal("||", ((ICsvDataReaderDialectMetadata)reader).DelimiterText);
        Assert.Equal(2, reader.FieldCount);
        Assert.True(await reader.ReadAsync());
        Assert.Equal("Alpha", reader.GetString(0));
        Assert.Equal("7", reader.GetString(1));
    }

    [Fact]
    public async Task Compressed_Path_Enforces_Compressed_Input_Limit_And_Releases_The_File()
    {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-csv-incremental-" + Guid.NewGuid().ToString("N") + ".csv.gz");
        try
        {
            await CsvDocument.Parse("Name\nAlpha\n").SaveAsync(path, new CsvSaveOptions { CompressionType = CsvCompressionType.GZip });
            await Assert.ThrowsAsync<InvalidDataException>(() => CsvDocument.OpenDataReaderAsync(path, new CsvLoadOptions { MaxInputBytes = 2 }));
            using var exclusive = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        }
        finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Synchronous_Failed_Advance_Hides_Previous_Row_And_Ends_Reader(bool mismatch)
    {
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(mismatch ? "Name\nAlpha\nBeta,Extra\nGamma\n" : "Name\nAlpha\n\"unfinished"));
        using var reader = await CsvDocument.OpenDataReaderAsync(input, new CsvLoadOptions
        {
            QuoteParsingMode = CsvQuoteParsingMode.Strict, ColumnCountMismatchPolicy = CsvColumnCountMismatchPolicy.Strict
        });
        Assert.True(await reader.ReadAsync());
        Assert.ThrowsAny<CsvException>(() => reader.Read());
        Assert.Equal(0, ((ICsvDataReaderPositionMetadata)reader).RecordNumber);
        Assert.Throws<InvalidOperationException>(() => reader.GetString(0));
        await Assert.ThrowsAsync<InvalidOperationException>(() => reader.ReadAsync());
    }

    [Fact]
    public async Task Cancellation_From_Progress_Callback_Ends_The_Advance()
    {
        using var cancellation = new CancellationTokenSource();
        using var input = new MemoryStream(Encoding.UTF8.GetBytes("Name\nAlpha\nBeta\n"));
        using var reader = await CsvDocument.OpenDataReaderAsync(input, new CsvLoadOptions
        {
            ProgressReportInterval = 1,
            ProgressCallback = progress => { if (progress.RecordsRead == 2) cancellation.Cancel(); }
        });
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => reader.ReadAsync(cancellation.Token));
        Assert.Equal(0, ((ICsvDataReaderPositionMetadata)reader).RecordNumber);
        await Assert.ThrowsAsync<InvalidOperationException>(() => reader.ReadAsync());
    }

    internal sealed class AsyncInput : Stream
    {
        private readonly byte[] _bytes;
        private readonly int _chunk;
        private readonly bool _block;
        private int _position;
        private bool _disposed;
        internal int BytesRead => _position;
        internal int AsyncReads { get; private set; }
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
            AsyncReads++;
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
