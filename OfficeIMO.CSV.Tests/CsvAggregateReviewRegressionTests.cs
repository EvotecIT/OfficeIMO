#if NET8_0_OR_GREATER
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.CSV;
using Xunit;
using AsyncInput = OfficeIMO.CSV.Tests.CsvIncrementalAsyncReaderTests.AsyncInput;

namespace OfficeIMO.CSV.Tests;

public sealed class CsvAggregateReviewRegressionTests {
    public enum Route { TextBatch, TextPartition, TextGeneral, TextSchema, AsyncSequential, AsyncParallel }

    [Theory]
    [InlineData(Route.TextBatch)]
    [InlineData(Route.TextPartition)]
    [InlineData(Route.TextGeneral)]
    [InlineData(Route.TextSchema)]
    [InlineData(Route.AsyncSequential)]
    [InlineData(Route.AsyncParallel)]
    public async Task TypedNullMarkersAreRejectedWithoutErasingRawTextAcrossReaderRoutes(Route route) {
        string[] markers = { "-1", "0", "1", "1.25", "NaN", "2024-01-02", "03366c08-b412-48af-9106-386b7da35f77" };
        var failures = new List<string>();
        for (int primitive = 0; primitive < markers.Length; primitive++) {
            string marker = markers[primitive];
            int getter = primitive, rows = 0;
            string text = "Value\n" + string.Concat(Enumerable.Repeat(marker + "\n", 8));
            var load = new CsvLoadOptions { NullValue = marker, MaxFieldLength = route == Route.TextGeneral ? 100 : null };
            var schema = route == Route.TextSchema ? new CsvSchemaBuilder().Column("Value").AsString().Done().Build() : null;
            var reader = schema is null ? null : new CsvDataReaderOptions { Schema = schema };
            var parallel = new ParallelRowMappingOptions {
                MaxDegreeOfParallelism = route is Route.TextBatch or Route.AsyncSequential ? 1 : 3,
                BatchSize = 1,
            };
            CsvRecordAccumulatorBuilder<List<string>> builder = _ => (ref List<string> errors, CsvRecord row) => {
                Assert.True(row.IsNull(0));
                Assert.False(row.IsMissing(0));
                Assert.Equal(marker, row.GetSpan(0).ToString());
                Assert.Equal(marker, row.GetString(0));
                string? failure = ReadNullPrimitive(row, getter);
                if (failure is not null) errors.Add(marker + ": " + failure);
                Interlocked.Increment(ref rows);
            };
            List<string> errors;
            if (route is Route.AsyncSequential or Route.AsyncParallel) {
                using var input = new AsyncInput(Encoding.UTF8.GetBytes(text), 1);
                errors = await CsvDocument.AggregateRowsAsParallelAsync(input, static () => new List<string>(),
                    builder, MergeErrors, load, reader, parallel);
                Assert.True(input.CanRead);
            }
            else errors = CsvDocument.AggregateTextRowsAsParallel(text, static () => new List<string>(),
                builder, MergeErrors, load, reader, parallel);
            Assert.Equal(8, rows);
            failures.AddRange(errors);
        }
        Assert.Empty(failures);
    }

    [Theory]
    [InlineData(Route.TextSchema)]
    [InlineData(Route.TextBatch)]
    [InlineData(Route.TextPartition)]
    [InlineData(Route.AsyncParallel)]
    public async Task SchemaDefaultsKeepCanonicalTypedSemanticsAndOriginalMarkerText(Route route) {
        const string text = "Value\n-1\n-1\n-1\n-1\n";
        var load = new CsvLoadOptions { NullValue = "-1" };
        var reader = new CsvDataReaderOptions {
            Schema = new CsvSchemaBuilder().Column("Value").AsInt32().WithDefault(42).Done().Build(),
        };
        CsvRecordAccumulatorBuilder<int> builder = _ => (ref int sum, CsvRecord row) => sum += ReadDefault(row);
        int result;
        if (route == Route.AsyncParallel) {
            using var input = new AsyncInput(Encoding.UTF8.GetBytes(text), 1);
            result = await CsvDocument.AggregateRowsAsParallelAsync<int>(input, static () => 0,
                builder, static (left, right) => left + right, load, reader, Options());
            Assert.True(input.CanRead);
        }
        else if (route is Route.TextBatch or Route.TextPartition)
            result = CsvDocument.ReadTextRowsAsParallel<int>(text, _ => row => ReadDefault(row), load, reader,
                new ParallelRowMappingOptions { MaxDegreeOfParallelism = route == Route.TextBatch ? 1 : 3, BatchSize = 1 }).Sum();
        else result = CsvDocument.AggregateTextRowsAsParallel<int>(text, static () => 0,
            builder, static (left, right) => left + right, load, reader, Options());
        Assert.Equal(168, result);
    }

    private static int ReadDefault(CsvRecord row) {
        Assert.Equal("-1", row.GetSpan(0).ToString());
        Assert.Equal("-1", row.GetString(0));
        Assert.False(row.IsNull(0));
        return row.GetInt32(0);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(3)]
    public void RecordProjectionKeepsCustomSchemaConversionAndRawText(int degree) {
        var reader = new CsvDataReaderOptions {
            Schema = new CsvSchemaBuilder().Column("Value").AsInt32()
                .ConvertUsing(value => Convert.ToInt32(value) + 100).Done().Build(),
        };
        int[] values = CsvDocument.ReadTextRowsAsParallel<int>("Value\n001\n002\n003\n004\n",
            _ => row => {
                Assert.Equal(3, row.GetSpan(0).Length);
                Assert.Equal(row.GetString(0), row.GetSpan(0).ToString());
                return row.GetInt32(0);
            }, readerOptions: reader,
            parallelOptions: new ParallelRowMappingOptions { MaxDegreeOfParallelism = degree, BatchSize = 1 }).ToArray();
        Assert.Equal(new[] { 101, 102, 103, 104 }, values);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task OriginalCallbackAndFactoryCancellationSurvivesBlockedAsyncProducer(bool factory) {
        using var input = new AsyncInput(Encoding.UTF8.GetBytes("Id\n1\n"), 1, blockAtEnd: true);
        using var operation = new CancellationTokenSource();
        using var load = new CancellationTokenSource();
        using var third = new CancellationTokenSource();
        var expected = new OperationCanceledException("original worker cancellation", new FormatException("inner"), third.Token);
        int factories = 0, active = 0;
        int Create() {
            if (Interlocked.Increment(ref factories) > 1 && factory) {
                Assert.True(input.Blocked.Task.Wait(TimeSpan.FromSeconds(10)));
                throw expected;
            }
            return 0;
        }
        OperationCanceledException actual = await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            CsvDocument.AggregateRowsAsParallelAsync<int>(input, Create,
                _ => (ref int rows, CsvRecord row) => {
                    Interlocked.Increment(ref active);
                    try {
                        Assert.True(input.Blocked.Task.Wait(TimeSpan.FromSeconds(10)));
                        throw expected;
                    }
                    finally { Interlocked.Decrement(ref active); }
                }, static (left, right) => left + right,
                loadOptions: new CsvLoadOptions { CancellationToken = load.Token },
                parallelOptions: Options(), cancellationToken: operation.Token).WaitAsync(TimeSpan.FromSeconds(10)));
        Assert.Same(expected, actual);
        Assert.Equal(third.Token, actual.CancellationToken);
        Assert.False(operation.IsCancellationRequested);
        Assert.False(load.IsCancellationRequested);
        Assert.True(input.Blocked.Task.IsCompleted);
        Assert.Equal(0, Volatile.Read(ref active));
        Assert.True(input.CanRead);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OriginalCallbackAndFactoryCancellationSurvivesSiblingTextWorkerCancellation(bool factory) {
        using var third = new CancellationTokenSource();
        var expected = new OperationCanceledException("original text worker cancellation", new FormatException("inner"), third.Token);
        int factories = 0, active = 0;
        string text = "Id\n" + string.Join("\n", Enumerable.Range(1, 64)) + "\n";
        OperationCanceledException actual = Assert.ThrowsAny<OperationCanceledException>(() =>
            CsvDocument.AggregateTextRowsAsParallel<int>(text,
                () => { if (Interlocked.Increment(ref factories) > 1 && factory) throw expected; return 0; },
                _ => (ref int rows, CsvRecord row) => {
                    Interlocked.Increment(ref active);
                    try { if (row.GetInt32(0) == 1) throw expected; rows++; }
                    finally { Interlocked.Decrement(ref active); }
                }, static (left, right) => left + right, parallelOptions: Options()));
        Assert.Same(expected, actual);
        Assert.Equal(third.Token, actual.CancellationToken);
        Assert.Equal(0, Volatile.Read(ref active));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task MergeCancellationKeepsOriginalExceptionAndJoinsStartedWorkers(bool asynchronous) {
        using var third = new CancellationTokenSource();
        var expected = new OperationCanceledException("original merge cancellation", new FormatException("inner"), third.Token);
        int active = 0;
        string text = "Id\n" + string.Join("\n", Enumerable.Range(1, 32)) + "\n";
        CsvRecordAccumulatorBuilder<int> builder = _ => (ref int rows, CsvRecord row) => {
            Interlocked.Increment(ref active);
            try { rows++; } finally { Interlocked.Decrement(ref active); }
        };
        OperationCanceledException actual;
        if (asynchronous) {
            using var input = new AsyncInput(Encoding.UTF8.GetBytes(text), 1);
            actual = await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                CsvDocument.AggregateRowsAsParallelAsync<int>(input, static () => 0, builder,
                    (left, right) => throw expected, parallelOptions: Options()));
            Assert.True(input.CanRead);
        }
        else actual = Assert.ThrowsAny<OperationCanceledException>(() =>
            CsvDocument.AggregateTextRowsAsParallel<int>(text, static () => 0, builder,
                (left, right) => throw expected, parallelOptions: Options()));
        Assert.Same(expected, actual);
        Assert.Equal(third.Token, actual.CancellationToken);
        Assert.Equal(0, Volatile.Read(ref active));
    }

    private static string? ReadNullPrimitive(CsvRecord row, int getter) {
        try {
            switch (getter) {
                case 0: row.GetInt32(0); break;
                case 1: row.GetInt64(0); break;
                case 2: row.GetBoolean(0); break;
                case 3: row.GetDecimal(0); break;
                case 4: row.GetDouble(0); break;
                case 5: row.GetDateTime(0); break;
                case 6: row.GetGuid(0); break;
            }
        }
        catch (InvalidCastException) { return null; }
        catch (Exception error) { return error.GetType().Name; }
        return "returned a primitive value for a null field";
    }

    private static List<string> MergeErrors(List<string> left, List<string> right) { left.AddRange(right); return left; }
    private static ParallelRowMappingOptions Options() => new() { MaxDegreeOfParallelism = 3, BatchSize = 1 };
}
#endif
