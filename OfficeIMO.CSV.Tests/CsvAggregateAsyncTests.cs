#if NET8_0_OR_GREATER
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.CSV;
using Xunit;
using AsyncInput = OfficeIMO.CSV.Tests.CsvIncrementalAsyncReaderTests.AsyncInput;

namespace OfficeIMO.CSV.Tests;

public sealed class CsvAggregateAsyncTests {
    [Theory]
    [InlineData(1)]
    [InlineData(3)]
    [InlineData(4096)]
    public async Task ShortReadsPreserveEveryBorrowedFieldAndOrderedIndependentState(int chunk) {
        Guid guid = Guid.Parse("03366c08-b412-48af-9106-386b7da35f77");
        string text = "Id,Name,Money,When,Flag,Key\r\n" + string.Concat(Enumerable.Range(1, 32).Select(
            id => $"{id},\"Żółć 😀 {id}\n\"\"quoted\"\"\",\"1,25\",02/01/2024,1,{guid}\r\n"));
        using var input = new AsyncInput(Encoding.UTF8.GetBytes(text), chunk);
        var created = new ConcurrentBag<AsyncState>();
        int builders = 0;
        AsyncState result = await CsvDocument.AggregateRowsAsParallelAsync(
            input, () => { var state = new AsyncState(); created.Add(state); return state; },
            header => {
                builders++;
                int id = header.GetOrdinal("Id"), name = header.GetOrdinal("Name");
                return (ref AsyncState state, CsvRecord row) => {
                    int value = row.GetInt32(id);
                    Assert.Equal(6, row.FieldCount);
                    Assert.Equal($"Żółć 😀 {value}\n\"quoted\"", row.GetSpan(name).ToString());
                    Assert.Equal(row.GetString(name), row.GetSpan(name).ToString());
                    Assert.Equal(1.25m, row.GetDecimal(2));
                    Assert.Equal(1.25, row.GetDouble(2));
                    Assert.Equal(new DateTime(2024, 1, 2), row.GetDateTime(3));
                    Assert.True(row.GetBoolean(4));
                    Assert.Equal(guid, row.GetGuid(5));
                    state.Ids.Add(value);
                    state.Sum += row.GetInt64(id);
                };
            }, static (left, right) => { Assert.NotSame(left, right); left.Ids.AddRange(right.Ids); left.Sum += right.Sum; return left; },
            loadOptions: new CsvLoadOptions {
                Culture = CultureInfo.GetCultureInfo("pl-PL"), DateTimeFormats = new[] { "dd/MM/yyyy" },
            }, parallelOptions: Options(3, 3));
        Assert.Equal(1, builders);
        Assert.Equal(Enumerable.Range(1, 32), result.Ids);
        Assert.Equal(528, result.Sum);
        Assert.True(created.Count > 1);
        Assert.True(input.AsyncReads > 0);
        Assert.True(input.CanRead);
    }

    [Fact]
    public async Task DetachedRecordsKeepNullMarkerMissingEmptyAndStaticFieldsDistinct() {
        using var input = new AsyncInput(Encoding.UTF8.GetBytes("Id,Value\n1,NULL\n2\n3,\n"), 1);
        Counts counts = await CsvDocument.AggregateRowsAsParallelAsync(
            input, static () => new Counts(),
            header => {
                Assert.Equal(4, header.Count);
                return (ref Counts state, CsvRecord row) => {
                    int id = row.GetInt32(0);
                    Assert.Equal(id == 1, row.IsNull(1));
                    Assert.Equal(id == 2, row.IsMissing(1));
                    Assert.Equal(id == 1 ? "NULL" : string.Empty, row.GetString(1));
                    Assert.Equal("7", row.GetSpan(2).ToString());
                    Assert.Equal(7, row.GetInt32(2));
                    Assert.False(row.IsMissing(2));
                    Assert.True(row.IsNull(3));
                    Assert.False(row.IsMissing(3));
                    Assert.True(row.GetSpan(3).IsEmpty);
                    state.Rows++;
                    if (row.IsNull(1)) state.Nulls++;
                    if (row.IsMissing(1)) state.Missing++;
                };
            }, static (left, right) => new Counts { Rows = left.Rows + right.Rows, Nulls = left.Nulls + right.Nulls, Missing = left.Missing + right.Missing },
            loadOptions: new CsvLoadOptions {
                NullValue = "NULL", StaticColumns = new Dictionary<string, object?> { ["Batch"] = 7, ["Empty"] = null },
            }, parallelOptions: Options(3, 1));
        Assert.Equal(3, counts.Rows);
        Assert.Equal(1, counts.Nulls);
        Assert.Equal(1, counts.Missing);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ExplicitAndInferredSchemaUseNativeAsyncFallbackWithoutLosingSourceText(bool inferred) {
        using var input = new AsyncInput(Encoding.UTF8.GetBytes("Id\n 001 \n 002 \n"), 1);
        int factories = 0, merges = 0;
        var schema = new CsvSchemaBuilder().Column("Id").AsInt32()
            .ConvertUsing(value => Convert.ToInt32(value, CultureInfo.InvariantCulture) + 100).Done().Build();
        int sum = await CsvDocument.AggregateRowsAsParallelAsync<int>(
            input, () => { factories++; return 0; },
            _ => (ref int total, CsvRecord row) => { Assert.Equal(3, row.GetSpan(0).Length); total += row.GetInt32(0); },
            (left, right) => { merges++; return left + right; },
            loadOptions: new CsvLoadOptions { TrimWhitespace = true, MaxFieldLength = 20 },
            readerOptions: new CsvDataReaderOptions { Schema = inferred ? null : schema, InferSchema = inferred, SchemaSampleSize = 2 },
            parallelOptions: Options(4, 1));
        Assert.Equal(inferred ? 3 : 203, sum);
        Assert.Equal(1, factories);
        Assert.Equal(0, merges);
        Assert.True(input.AsyncReads > 0);
        Assert.True(input.CanRead);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task BoundedReadAheadAndCancellationJoinAnAlreadyRunningCallback(bool loadToken) {
        byte[] bytes = Encoding.UTF8.GetBytes("Id\n" + string.Concat(Enumerable.Repeat("1\n", 20_000)));
        using var input = new AsyncInput(bytes, 1, blockAtEnd: true);
        using var cancel = new CancellationTokenSource();
        using var release = new ManualResetEventSlim();
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        int active = 0;
        Task<int> work = CsvDocument.AggregateRowsAsParallelAsync<int>(
            input, static () => 0,
            _ => (ref int rows, CsvRecord row) => {
                Interlocked.Increment(ref active);
                try { entered.TrySetResult(); release.Wait(); rows += row.GetInt32(0); }
                finally { Interlocked.Decrement(ref active); }
            }, static (left, right) => left + right,
            loadOptions: new CsvLoadOptions { CancellationToken = loadToken ? cancel.Token : default },
            parallelOptions: Options(2, 2), cancellationToken: loadToken ? default : cancel.Token);
        try {
            await entered.Task.WaitAsync(TimeSpan.FromSeconds(10));
            Assert.True(input.BytesRead < bytes.Length);
            Assert.False(input.Blocked.Task.IsCompleted);
            cancel.Cancel();
            Assert.False(work.IsCompleted);
        }
        finally { release.Set(); }
        OperationCanceledException error = await Assert.ThrowsAnyAsync<OperationCanceledException>(() => work.WaitAsync(TimeSpan.FromSeconds(10)));
        Assert.Equal(cancel.Token, error.CancellationToken);
        Assert.Equal(0, Volatile.Read(ref active));
        Assert.True(input.CanRead);
    }

    [Fact]
    public async Task WorkerFailureInterruptsBlockedProducerAndPreservesOriginalException() {
        using var input = new AsyncInput(Encoding.UTF8.GetBytes("Id\n1\n"), 1, blockAtEnd: true);
        var expected = new FormatException("worker failure");
        int active = 0, merges = 0;
        FormatException actual = await Assert.ThrowsAsync<FormatException>(() =>
            CsvDocument.AggregateRowsAsParallelAsync<int>(
                input, static () => 0,
                _ => (ref int rows, CsvRecord row) => {
                    Interlocked.Increment(ref active);
                    try { throw expected; } finally { Interlocked.Decrement(ref active); }
                }, (left, right) => { merges++; return left + right; }, parallelOptions: Options(3, 1)).WaitAsync(TimeSpan.FromSeconds(10)));
        Assert.Same(expected, actual);
        Assert.Equal(0, merges);
        Assert.Equal(0, Volatile.Read(ref active));
        Assert.True(input.CanRead);
    }

    [Fact]
    public async Task MergeFailureJoinsWorkersWithoutDisposingCallerOwnedStates() {
        using var input = new AsyncInput(Encoding.UTF8.GetBytes("Id\n" + string.Join("\n", Enumerable.Range(1, 32)) + "\n"), 1);
        var created = new ConcurrentBag<AsyncState>();
        var expected = new InvalidOperationException("merge failure");
        int active = 0, merges = 0;
        InvalidOperationException actual = await Assert.ThrowsAsync<InvalidOperationException>(() =>
            CsvDocument.AggregateRowsAsParallelAsync(input,
                () => { var state = new AsyncState(); created.Add(state); return state; },
                _ => (ref AsyncState state, CsvRecord row) => {
                    Interlocked.Increment(ref active);
                    try { state.Ids.Add(row.GetInt32(0)); }
                    finally { Interlocked.Decrement(ref active); }
                }, (left, right) => { merges++; throw expected; }, parallelOptions: Options(3, 1)));
        Assert.Same(expected, actual);
        Assert.Equal(1, merges);
        Assert.Equal(0, Volatile.Read(ref active));
        Assert.All(created, state => Assert.False(state.Disposed));
        Assert.True(input.CanRead);
    }

    [Fact]
    public async Task CancellationRequestedByFinalMergeIsObservedBeforeReturning() {
        using var input = new AsyncInput(Encoding.UTF8.GetBytes("Id\n" + string.Join("\n", Enumerable.Range(1, 32)) + "\n"), 1);
        using var cancel = new CancellationTokenSource();
        OperationCanceledException error = await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            CsvDocument.AggregateRowsAsParallelAsync<int>(input, static () => 0,
                _ => (ref int rows, CsvRecord record) => rows++,
                (left, right) => {
                    int rows = left + right;
                    if (rows == 32) cancel.Cancel();
                    return rows;
                }, parallelOptions: Options(3, 4), cancellationToken: cancel.Token));
        Assert.Equal(cancel.Token, error.CancellationToken);
        Assert.True(input.CanRead);
    }

    [Fact]
    public async Task InputFieldAndWidthLimitsEndTheOperationAndPreserveCallerStream() {
        await AssertLimit<InvalidDataException>("V\nlong\n", new CsvLoadOptions { MaxInputBytes = 2 });
        await AssertLimit<CsvParseException>("V\nlong\n", new CsvLoadOptions { MaxFieldLength = 3 });
        await AssertLimit<CsvException>("A,B\n1,2\n3\n", new CsvLoadOptions { ColumnCountMismatchPolicy = CsvColumnCountMismatchPolicy.Strict });
        await AssertLimit<CsvParseException>("Id,Name\n1,first\n2,\"unfinished", new CsvLoadOptions { QuoteParsingMode = CsvQuoteParsingMode.Strict });
    }

    [Fact]
    public async Task FileReaderClosesAfterSuccessAndCallbackFailure() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-csv-aggregate-" + Guid.NewGuid().ToString("N") + ".csv");
        try {
            await File.WriteAllTextAsync(path, "Id\n1\n2\n");
            int result = await CsvDocument.AggregateRowsAsParallelAsync<int>(path, static () => 0,
                _ => (ref int sum, CsvRecord row) => sum += row.GetInt32(0), static (left, right) => left + right, parallelOptions: Options(2, 1));
            Assert.Equal(3, result);
            using (File.Open(path, FileMode.Open, FileAccess.Read, FileShare.None)) { }
            var expected = new FormatException("file callback failure");
            Assert.Same(expected, await Assert.ThrowsAsync<FormatException>(() =>
                CsvDocument.AggregateRowsAsParallelAsync<int>(path, static () => 0,
                    _ => (ref int sum, CsvRecord row) => throw expected, static (left, right) => left + right, parallelOptions: Options(2, 1))));
            using (File.Open(path, FileMode.Open, FileAccess.Read, FileShare.None)) { }
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public async Task StreamStartsAtCurrentPositionAndKeepsCallerOwnership() {
        byte[] prefix = Encoding.UTF8.GetBytes("ignored\n");
        using var input = new MemoryStream(prefix.Concat(Encoding.UTF8.GetBytes("Id\n1\n2\n")).ToArray());
        input.Position = prefix.Length;
        int sum = await CsvDocument.AggregateRowsAsParallelAsync<int>(input, static () => 0,
            _ => (ref int state, CsvRecord row) => state += row.GetInt32(0), static (left, right) => left + right, parallelOptions: Options(2, 1));
        Assert.Equal(3, sum);
        Assert.Equal(input.Length, input.Position);
        Assert.True(input.CanRead);
    }

    [Fact]
    public async Task HeaderlessCompressedBomInputRetainsItsFirstRecord() {
        using var encoded = new MemoryStream();
        await CsvDocument.Parse("Id,Name\n1,Żółć\n2,last\n").SaveAsync(encoded,
            new CsvSaveOptions { IncludeHeader = false, Encoding = new UnicodeEncoding(false, true), CompressionType = CsvCompressionType.GZip });
        using var input = new AsyncInput(encoded.ToArray(), 2);
        AsyncState result = await CsvDocument.AggregateRowsAsParallelAsync(input, static () => new AsyncState(),
            header => {
                Assert.Equal("Column1", header[0]);
                return (ref AsyncState state, CsvRecord row) => {
                    int id = row.GetInt32(0);
                    Assert.Equal(id == 1 ? "Żółć" : "last", row.GetSpan(1).ToString());
                    state.Ids.Add(id);
                };
            }, static (left, right) => { left.Ids.AddRange(right.Ids); return left; },
            loadOptions: new CsvLoadOptions { HasHeaderRow = false, CompressionType = CsvCompressionType.GZip }, parallelOptions: Options(2, 1));
        Assert.Equal(new[] { 1, 2 }, result.Ids);
        Assert.True(input.CanRead);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PreCancellationReadsNothingAndPreservesOriginalToken(bool loadToken) {
        using var input = new AsyncInput(Encoding.UTF8.GetBytes("Id\n1\n"), 1);
        using var cancel = new CancellationTokenSource();
        cancel.Cancel();
        int factories = 0, builders = 0;
        OperationCanceledException error = await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            CsvDocument.AggregateRowsAsParallelAsync<int>(input, () => { factories++; return 0; },
                _ => { builders++; return (ref int sum, CsvRecord row) => sum++; }, static (left, right) => left + right,
                loadOptions: new CsvLoadOptions { CancellationToken = loadToken ? cancel.Token : default },
                cancellationToken: loadToken ? default : cancel.Token));
        Assert.Equal(cancel.Token, error.CancellationToken);
        Assert.Equal(0, input.AsyncReads);
        Assert.Equal(0, factories);
        Assert.Equal(0, builders);
        Assert.True(input.CanRead);
    }

    [Theory]
    [InlineData("")]
    [InlineData("Id\n")]
    public async Task EmptyInputReturnsNeutralStateWithoutRecordOrMergeCalls(string text) {
        using var input = new AsyncInput(Encoding.UTF8.GetBytes(text), 1);
        int creates = 0;
        int result = await CsvDocument.AggregateRowsAsParallelAsync<int>(input, () => { creates++; return 0; },
            _ => (ref int sum, CsvRecord row) => throw new InvalidOperationException("unexpected record"),
            static (left, right) => throw new InvalidOperationException("unexpected merge"), parallelOptions: Options(3, 2));
        Assert.Equal(0, result);
        Assert.Equal(1, creates);
        Assert.True(input.CanRead);
    }

    private static async Task AssertLimit<TException>(string text, CsvLoadOptions options) where TException : Exception {
        using var input = new AsyncInput(Encoding.UTF8.GetBytes(text), 1);
        int active = 0;
        await Assert.ThrowsAnyAsync<TException>(() => CsvDocument.AggregateRowsAsParallelAsync<int>(input, static () => 0,
            _ => (ref int rows, CsvRecord row) => {
                Interlocked.Increment(ref active);
                try { rows++; } finally { Interlocked.Decrement(ref active); }
            }, static (left, right) => left + right, loadOptions: options, parallelOptions: Options(3, 1)));
        Assert.Equal(0, Volatile.Read(ref active));
        Assert.True(input.CanRead);
    }

    private static ParallelRowMappingOptions Options(int degree, int size) => new() { MaxDegreeOfParallelism = degree, BatchSize = size };
    private struct Counts { internal int Rows, Nulls, Missing; }
    private sealed class AsyncState : IDisposable {
        internal readonly List<int> Ids = new();
        internal long Sum;
        internal bool Disposed;
        public void Dispose() => Disposed = true;
    }
}
#endif
