#if NET8_0_OR_GREATER
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.CSV;
using Xunit;
using AsyncInput = OfficeIMO.CSV.Tests.CsvIncrementalAsyncReaderTests.AsyncInput;

namespace OfficeIMO.CSV.Tests;

public sealed class CsvAggregateBatchReuseTests {
    [Fact]
    public async Task SuccessiveMixedBatchesKeepRetainedStringsIndependentStatesAndSourceOrder() {
        const int rows = 19, batchSize = 3;
        var expected = new List<(int Id, string Name, string Value, string Tail)>();
        var text = new StringBuilder("Id,Name,Value,Tail\r\n");
        for (int id = 1; id <= rows; id++) {
            string name = id == 1 ? new string('x', 70_000) + ",Żółć 😀\n\"quoted\""
                : id % 5 == 0 ? "" : id % 3 == 2 ? $"quoted, \"name\" {id}\nŻółć 😀" : $"raw Żółć 😀 {id}";
            bool missing = id % 4 == 3;
            string value = missing || id % 4 == 1 ? "" : id % 4 == 0 ? "NULL" : $"value{id}";
            string tail = !missing && id % 5 == 0 ? $"tail{id}" : "";
            expected.Add((id, name, value, tail));
            text.Append(id).Append(',').Append(name.IndexOfAny(new[] { ',', '\r', '\n', '"' }) < 0 ? " " + name + " " : Encode(name));
            if (!missing) text.Append(',').Append(value).Append(',').Append(tail);
            if (id % 6 == 0) text.Append(",ignored");
            text.Append("\r\n");
        }

        using var input = new AsyncInput(Encoding.UTF8.GetBytes(text.ToString()), 4096);
        using var secondBatchFinished = new ManualResetEventSlim();
        var completionOrder = new ConcurrentQueue<int>();
        var created = new ConcurrentBag<BatchState>();
        var merged = new List<BatchState>();
        BatchState result = await CsvDocument.AggregateRowsAsParallelAsync(
            input, () => { var state = new BatchState(); created.Add(state); return state; },
            header => {
                Assert.Equal(4, header.Count);
                return (ref BatchState state, CsvRecord row) => {
                    int id = row.GetInt32(0);
                    var item = expected[id - 1];
                    Assert.Equal(4, row.FieldCount);
                    Assert.Equal((id - 1) % batchSize, state.Rows.Count);
                    Assert.Equal(item.Name, row.GetSpan(1).ToString());
                    Assert.Equal(item.Value, row.GetSpan(2).ToString());
                    Assert.Equal(item.Tail, row.GetSpan(3).ToString());
                    string name = row.GetString(1), value = row.GetString(2), tail = row.GetString(3);
                    Assert.Equal(item.Name, row.GetSpan(1).ToString());
                    Assert.Equal(item.Value, row.GetSpan(2).ToString());
                    Assert.Equal(item.Tail, row.GetSpan(3).ToString());
                    Assert.Equal(id % 4 == 3, row.IsMissing(2));
                    Assert.Equal(id % 4 == 3, row.IsMissing(3));
                    Assert.Equal(id % 4 == 0, row.IsNull(2));
                    Assert.False(row.IsNull(3));
                    state.Rows.Add((id, name, value, tail));
                    if (id == 3) Assert.True(secondBatchFinished.Wait(TimeSpan.FromSeconds(10)));
                    if (id % batchSize == 0) completionOrder.Enqueue(id / batchSize);
                    if (id == 6) secondBatchFinished.Set();
                };
            }, (left, right) => {
                Assert.NotSame(left, right);
                int firstId = merged.Count * batchSize + 1;
                Assert.Equal(Enumerable.Range(firstId, Math.Min(batchSize, rows - firstId + 1)), right.Rows.Select(row => row.Id));
                merged.Add(right);
                left.Rows.AddRange(right.Rows);
                return left;
            }, loadOptions: new CsvLoadOptions {
                TrimWhitespace = true, NullValue = "NULL",
                ColumnCountMismatchPolicy = CsvColumnCountMismatchPolicy.PadMissingFieldsAndIgnoreExtraFields,
            }, parallelOptions: new ParallelRowMappingOptions { MaxDegreeOfParallelism = 2, BatchSize = batchSize })
            .WaitAsync(TimeSpan.FromSeconds(20));

        Assert.Equal(new[] { 2, 1 }, completionOrder.Take(2));
        Assert.Equal(8, created.Count); // One neutral result and seven original input batches.
        Assert.Equal(7, merged.Count);
        Assert.Equal(expected, result.Rows);
        Assert.Equal(expected, merged.SelectMany(state => state.Rows));
        Assert.True(input.AsyncReads > 0);
        Assert.True(input.CanRead);
    }

    [Fact]
    public async Task EachDispatchCapturesMergeContextWithoutRetainingPriorWorkerChanges() {
        const int degree = 2, rows = 7;
        var ambient = new AsyncLocal<int> { Value = 0 };
        using var input = new AsyncInput(Encoding.UTF8.GetBytes("Id\n" + string.Join("\n", Enumerable.Range(1, rows)) + "\n"), 1);
        ContextState result = await CsvDocument.AggregateRowsAsParallelAsync(
            input, () => new ContextState(ambient.Value),
            _ => (ref ContextState state, CsvRecord row) => {
                int id = row.GetInt32(0);
                Assert.Equal(Math.Max(0, id - degree), state.FactoryContext);
                Assert.Equal(state.FactoryContext, ambient.Value);
                state.Ids.Add(id);
                ambient.Value = -id;
            }, (left, right) => {
                int id = Assert.Single(right.Ids);
                Assert.Equal(id - 1, ambient.Value);
                ambient.Value = id;
                left.Ids.Add(id);
                return left;
            }, parallelOptions: new ParallelRowMappingOptions { MaxDegreeOfParallelism = degree, BatchSize = 1 });
        Assert.Equal(Enumerable.Range(1, rows), result.Ids);
        Assert.Equal(0, ambient.Value);
        Assert.True(input.AsyncReads > 0);
        Assert.True(input.CanRead);
    }

    [Fact]
    public async Task CancellationJoinsReusedSlotBeforeReleasingItsCallbackText() {
        using var input = new AsyncInput(Encoding.UTF8.GetBytes("Id,Name\n" + string.Concat(Enumerable.Range(1, 7).Select(
            id => $"{id},row Żółć 😀 {id}\n"))), 1);
        using var cancel = new CancellationTokenSource();
        using var release = new ManualResetEventSlim();
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        int active = 0, merges = 0;
        Exception? callbackFailure = null;
        string? retained = null;
        Task<int> work = CsvDocument.AggregateRowsAsParallelAsync<int>(
            input, static () => 0,
            _ => (ref int total, CsvRecord row) => {
                Interlocked.Increment(ref active);
                try {
                    int id = row.GetInt32(0);
                    if (id == 3) {
                        Assert.True(Volatile.Read(ref merges) >= 1);
                        ReadOnlySpan<char> name = row.GetSpan(1);
                        entered.TrySetResult();
                        release.Wait();
                        Assert.Equal("row Żółć 😀 3", name.ToString());
                        retained = row.GetString(1);
                    }
                    total += id;
                }
                catch (Exception exception) { Interlocked.CompareExchange(ref callbackFailure, exception, null); throw; }
                finally { Interlocked.Decrement(ref active); }
            }, (left, right) => { Interlocked.Increment(ref merges); return left + right; },
            parallelOptions: new ParallelRowMappingOptions { MaxDegreeOfParallelism = 2, BatchSize = 1 }, cancellationToken: cancel.Token);
        try {
            await entered.Task.WaitAsync(TimeSpan.FromSeconds(10));
            cancel.Cancel();
            Assert.False(work.IsCompleted);
        }
        finally { release.Set(); }
        OperationCanceledException error = await Assert.ThrowsAnyAsync<OperationCanceledException>(() => work.WaitAsync(TimeSpan.FromSeconds(10)));
        Assert.Equal(cancel.Token, error.CancellationToken);
        Assert.Null(callbackFailure);
        Assert.Equal("row Żółć 😀 3", retained);
        Assert.Equal(0, Volatile.Read(ref active));
        Assert.True(input.CanRead);
    }

    private static string Encode(string value) => value.IndexOfAny(new[] { ',', '\r', '\n', '"' }) < 0
        ? value : "\"" + value.Replace("\"", "\"\"") + "\"";

    private sealed class BatchState {
        internal readonly List<(int Id, string Name, string Value, string Tail)> Rows = new();
    }

    private sealed class ContextState {
        internal ContextState(int factoryContext) => FactoryContext = factoryContext;
        internal int FactoryContext { get; }
        internal readonly List<int> Ids = new();
    }
}
#endif
