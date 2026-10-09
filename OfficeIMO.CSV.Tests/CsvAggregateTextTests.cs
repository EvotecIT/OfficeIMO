#if NET8_0_OR_GREATER
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using OfficeIMO.CSV;
using OfficeIMO.Data;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public sealed class CsvAggregateTextTests {
    [Theory]
    [InlineData(1)]
    [InlineData(4)]
    public void StructStateConsumesBorrowedQuotedUnicodeFieldsAndResolvesHeaderOnce(int degree) {
        string text = "Id,Name,Date\n" + string.Concat(Enumerable.Range(1, 32).Select(
            id => $"{id},\"Żółć 😀 {id}\n\"\"quoted\"\"\",2024-01-02T03:04:05.0000000\n"));
        int builders = 0;
        AggregateStats result = CsvDocument.AggregateTextRowsAsParallel(
            text, static () => new AggregateStats(),
            header => {
                builders++;
                int id = header.GetOrdinal("Id"), name = header.GetOrdinal("Name"), date = header.GetOrdinal("Date");
                return (ref AggregateStats state, CsvRecord row) => {
                    state.Rows++;
                    state.Sum += row.GetInt32(id);
                    state.Characters += row.GetSpan(name).Length;
                    Assert.Equal($"Żółć 😀 {row.GetInt32(id)}\n\"quoted\"", row.GetString(name));
                    Assert.Equal(new DateTime(2024, 1, 2, 3, 4, 5), row.GetDateTime(date));
                };
            }, MergeStats, parallelOptions: Options(degree));
        Assert.Equal(1, builders);
        Assert.Equal(32, result.Rows);
        Assert.Equal(528, result.Sum);
        Assert.Equal(Enumerable.Range(1, 32).Sum(id => $"Żółć 😀 {id}\n\"quoted\"".Length), result.Characters);
    }

    [Fact]
    public void NoncommutativeMergeRetainsSourceOrderAndIndependentCallerOwnedStates() {
        string text = "Id\n" + string.Join("\n", Enumerable.Range(1, 64)) + "\n";
        var created = new ConcurrentBag<OwnedState>();
        int caller = Environment.CurrentManagedThreadId;
        OwnedState result = CsvDocument.AggregateTextRowsAsParallel(
            text, () => { var state = new OwnedState(); created.Add(state); return state; },
            header => {
                Assert.Equal(caller, Environment.CurrentManagedThreadId);
                int id = header.GetOrdinal("Id");
                return (ref OwnedState state, CsvRecord row) => state.Ids.Add(row.GetInt32(id));
            },
            (left, right) => {
                Assert.Equal(caller, Environment.CurrentManagedThreadId);
                Assert.NotSame(left, right);
                left.Ids.AddRange(right.Ids);
                return left;
            }, parallelOptions: Options(4));
        Assert.Equal(Enumerable.Range(1, 64), result.Ids);
        Assert.True(created.Count > 1);
        Assert.All(created, state => Assert.False(state.Disposed));
        result.Dispose();
        Assert.True(result.Disposed);
    }

    [Fact]
    public void SchemaFallbackRetainsRawSpansAndTypedConverterSemantics() {
        CsvSchema schema = new CsvSchemaBuilder().Column("Id").AsInt32()
            .ConvertUsing(value => Convert.ToInt32(value, System.Globalization.CultureInfo.InvariantCulture) + 100).Done().Build();
        AggregateStats result = CsvDocument.AggregateTextRowsAsParallel(
            "Id\n 001 \n 002 \n", static () => new AggregateStats(),
            _ => (ref AggregateStats state, CsvRecord row) => {
                state.Rows++;
                state.Sum += row.GetInt32(0);
                state.Characters += row.GetSpan(0).Length;
                Assert.Equal(row.GetString(0), row.GetSpan(0).ToString());
            }, MergeStats, loadOptions: new CsvLoadOptions { TrimWhitespace = true },
            readerOptions: new CsvDataReaderOptions { Schema = schema }, parallelOptions: Options(4));
        Assert.Equal(2, result.Rows);
        Assert.Equal(203, result.Sum);
        Assert.Equal(6, result.Characters);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StaticColumnFallbackExposesEveryResolvedFieldToAggregateAndProjection(bool projection) {
        var load = new CsvLoadOptions {
            StaticColumns = new Dictionary<string, object?> {
                ["Batch"] = 7, ["Label"] = "Żółć", ["Empty"] = null,
            },
        };
        const string text = "Id\n1\n2\n";
        if (projection) {
            Assert.Equal(new[] { 8, 9 }, CsvDocument.ReadTextRowsAsParallel<int>(
                text, _ => row => ValidateStaticFields(row), load, parallelOptions: Options(4)));
            return;
        }
        int result = CsvDocument.AggregateTextRowsAsParallel<int>(
            text, static () => 0,
            header => {
                Assert.Equal(4, header.Count);
                return (ref int total, CsvRecord row) => total += ValidateStaticFields(row);
            }, static (left, right) => left + right,
            loadOptions: load, parallelOptions: Options(4));
        Assert.Equal(17, result);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InferredGeneralParserFallbackPreservesSourceTextNullAndMissingFields(bool projection) {
        const string text = "Id,Value\n001,NULL\n002\n003,\n";
        var load = new CsvLoadOptions { NullValue = "NULL", MaxFieldLength = 20 };
        var reader = new CsvDataReaderOptions { InferSchema = true, SchemaSampleSize = 2 };
        if (projection) {
            Assert.Equal(new[] { 1, 2, 3 }, CsvDocument.ReadTextRowsAsParallel<int>(
                text, _ => row => ValidateInferredFields(row), load, reader, Options(4)));
            return;
        }
        int sum = CsvDocument.AggregateTextRowsAsParallel<int>(
            text, static () => 0,
            _ => (ref int value, CsvRecord row) => value += ValidateInferredFields(row),
            static (left, right) => left + right, load, reader, Options(4));
        Assert.Equal(6, sum);
    }

    private static int ValidateInferredFields(CsvRecord row) {
        int id = row.GetInt32(0);
        Assert.Equal(id.ToString("D3"), row.GetSpan(0).ToString());
        Assert.Equal(id == 1, row.IsNull(1));
        Assert.Equal(id == 2, row.IsMissing(1));
        Assert.Equal(id == 1 ? "NULL" : string.Empty, row.GetString(1));
        return id;
    }

    private static int ValidateStaticFields(CsvRecord row) {
        Assert.Equal(4, row.FieldCount);
        Assert.Equal("7", row.GetString(1));
        Assert.Equal("7", row.GetSpan(1).ToString());
        Assert.Equal("Żółć", row.GetString(2));
        Assert.False(row.IsMissing(1));
        Assert.True(row.IsNull(3));
        Assert.False(row.IsMissing(3));
        Assert.True(row.GetSpan(3).IsEmpty);
        return row.GetInt32(0) + row.GetInt32(1);
    }

    [Fact]
    public void PartitionAndBatchRecordsKeepNullMissingAndEmptyDistinct() {
        foreach (int degree in new[] { 1, 4 }) {
            int[] result = CsvDocument.AggregateTextRowsAsParallel(
                "Id,Value\n1,NULL\n2\n3,\n", static () => new int[3],
                _ => (ref int[] counts, CsvRecord row) => {
                    if (row.IsNull(1)) counts[0]++;
                    else if (row.IsMissing(1)) counts[1]++;
                    else if (row.GetSpan(1).IsEmpty) counts[2]++;
                }, static (left, right) => {
                    for (int index = 0; index < left.Length; index++) left[index] += right[index];
                    return left;
                }, loadOptions: new CsvLoadOptions { NullValue = "NULL" },
                parallelOptions: new ParallelRowMappingOptions { MaxDegreeOfParallelism = degree, BatchSize = 1 });
            Assert.Equal(new[] { 1, 1, 1 }, result);
        }
    }

    [Fact]
    public void LongQuotedFallbackContinuesAfterPreparedBatchesWithoutLosingRows() {
        string longValue = new('x', 70_000);
        string text = "Value\n" + string.Concat(Enumerable.Repeat("short\n", 32)) + $"\"{longValue}\"\nlast\n";
        AggregateStats result = CsvDocument.AggregateTextRowsAsParallel(
            text, static () => new AggregateStats(),
            _ => (ref AggregateStats state, CsvRecord row) => { state.Rows++; state.Characters += row.GetSpan(0).Length; },
            MergeStats, parallelOptions: Options(4));
        Assert.Equal(34, result.Rows);
        Assert.Equal(70_164, result.Characters);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BothCancellationTokensStopBeforeBuildingHeaderOrState(bool loadToken) {
        using var cancel = new CancellationTokenSource();
        cancel.Cancel();
        int callbacks = 0;
        OperationCanceledException error = Assert.ThrowsAny<OperationCanceledException>(() =>
            CsvDocument.AggregateTextRowsAsParallel<int>(
                "Id\n1\n", () => { callbacks++; return 0; },
                _ => { callbacks++; return (ref int sum, CsvRecord row) => sum += row.GetInt32(0); },
                static (left, right) => left + right,
                loadOptions: new CsvLoadOptions { CancellationToken = loadToken ? cancel.Token : default },
                cancellationToken: loadToken ? default : cancel.Token));
        Assert.Equal(cancel.Token, error.CancellationToken);
        Assert.Equal(0, callbacks);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CallbackCancellationJoinsWorkersAndPreservesTheRequestedToken(bool loadToken) {
        using var cancel = new CancellationTokenSource();
        int active = 0;
        string text = "Id\n" + string.Join("\n", Enumerable.Range(1, 64)) + "\n";
        OperationCanceledException error = Assert.ThrowsAny<OperationCanceledException>(() =>
            CsvDocument.AggregateTextRowsAsParallel<int>(
                text, static () => 0,
                _ => (ref int sum, CsvRecord row) => {
                    Interlocked.Increment(ref active);
                    try { sum += row.GetInt32(0); cancel.Cancel(); }
                    finally { Interlocked.Decrement(ref active); }
                }, static (left, right) => left + right,
                loadOptions: new CsvLoadOptions { CancellationToken = loadToken ? cancel.Token : default },
                parallelOptions: Options(4), cancellationToken: loadToken ? default : cancel.Token));
        Assert.Equal(cancel.Token, error.CancellationToken);
        Assert.Equal(0, Volatile.Read(ref active));
    }

    [Fact]
    public void CallbackFailurePropagatesOriginalErrorAfterWorkersFinish() {
        var expected = new FormatException("record failure");
        int active = 0, merges = 0;
        string text = "Id\n" + string.Join("\n", Enumerable.Range(1, 64)) + "\n";
        FormatException actual = Assert.Throws<FormatException>(() => CsvDocument.AggregateTextRowsAsParallel<int>(
            text, static () => 0,
            _ => (ref int sum, CsvRecord row) => {
                Interlocked.Increment(ref active);
                try { if (row.GetInt32(0) == 1) throw expected; sum++; }
                finally { Interlocked.Decrement(ref active); }
            }, (left, right) => { merges++; return left + right; }, parallelOptions: Options(4)));
        Assert.Same(expected, actual);
        Assert.Equal(0, merges);
        Assert.Equal(0, Volatile.Read(ref active));
    }

    [Fact]
    public void ReaderLimitsAndStrictWidthErrorsRemainEnforced() {
        Assert.Throws<InvalidDataException>(() => Count("Value\nlong\n", new CsvLoadOptions { MaxInputBytes = 2 }));
        Assert.Throws<CsvParseException>(() => Count("V\nlong\n", new CsvLoadOptions { MaxFieldLength = 3 }));
        Assert.Throws<CsvException>(() => Count("A,B\n1,2\n3\n",
            new CsvLoadOptions { ColumnCountMismatchPolicy = CsvColumnCountMismatchPolicy.Strict }));
    }

    [Fact]
    public void MergeFailureDoesNotDisposeCallerStatesAndLeavesNoWorkerRunning() {
        var expected = new InvalidOperationException("merge failure");
        var created = new ConcurrentBag<OwnedState>();
        int active = 0, merges = 0;
        string text = "Id\n" + string.Join("\n", Enumerable.Range(1, 64)) + "\n";
        InvalidOperationException actual = Assert.Throws<InvalidOperationException>(() =>
            CsvDocument.AggregateTextRowsAsParallel(
                text, () => { var state = new OwnedState(); created.Add(state); return state; },
                _ => (ref OwnedState state, CsvRecord row) => {
                    Interlocked.Increment(ref active);
                    try { state.Ids.Add(row.GetInt32(0)); }
                    finally { Interlocked.Decrement(ref active); }
                }, (left, right) => { merges++; throw expected; }, parallelOptions: Options(4)));
        Assert.Same(expected, actual);
        Assert.Equal(1, merges);
        Assert.Equal(0, Volatile.Read(ref active));
        Assert.All(created, state => Assert.False(state.Disposed));
    }

    [Fact]
    public void CancellationRequestedByFinalMergeIsObservedBeforeReturning() {
        using var cancel = new CancellationTokenSource();
        string text = "Id\n" + string.Join("\n", Enumerable.Range(1, 32)) + "\n";
        OperationCanceledException error = Assert.ThrowsAny<OperationCanceledException>(() =>
            CsvDocument.AggregateTextRowsAsParallel<int>(
                text, static () => 0,
                _ => (ref int rows, CsvRecord row) => rows++,
                (left, right) => {
                    int rows = left + right;
                    if (rows == 32) cancel.Cancel();
                    return rows;
                }, parallelOptions: Options(4), cancellationToken: cancel.Token));
        Assert.Equal(cancel.Token, error.CancellationToken);
    }

    [Theory]
    [InlineData("")]
    [InlineData("Id\n")]
    public void EmptyInputReturnsNeutralStateWithoutRecordOrMergeCalls(string text) {
        int creates = 0;
        int result = CsvDocument.AggregateTextRowsAsParallel<int>(
            text, () => { creates++; return 0; },
            _ => (ref int value, CsvRecord row) => throw new InvalidOperationException("unexpected row"),
            static (left, right) => throw new InvalidOperationException("unexpected merge"), parallelOptions: Options(4));
        Assert.Equal(0, result);
        Assert.Equal(1, creates);
    }

    private static int Count(string text, CsvLoadOptions options) => CsvDocument.AggregateTextRowsAsParallel<int>(
        text, static () => 0, _ => (ref int count, CsvRecord row) => count++,
        static (left, right) => left + right, loadOptions: options, parallelOptions: Options(4));

    private static ParallelRowMappingOptions Options(int degree) => new() { MaxDegreeOfParallelism = degree, BatchSize = 4 };

    private static AggregateStats MergeStats(AggregateStats left, AggregateStats right) => new() {
        Rows = left.Rows + right.Rows, Sum = left.Sum + right.Sum, Characters = left.Characters + right.Characters,
    };

    private struct AggregateStats { internal int Rows, Sum, Characters; }
    private sealed class OwnedState : IDisposable {
        internal readonly List<int> Ids = new();
        internal bool Disposed;
        public void Dispose() => Disposed = true;
    }
}
#endif
