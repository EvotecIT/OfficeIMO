#if NET8_0_OR_GREATER
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Data;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public sealed class CsvAsyncMappingCancellationTests
{
    [Fact]
    public async Task RowsAsAsync_Automatic_DoesNotYieldMappedRowAfterCancellation()
    {
        using var cancellation = new CancellationTokenSource();
        using var reader = CreateCancelingReader(cancellation);

        await AssertCanceledBeforeYield(
            reader.RowsAsAsync<AsyncProbeRow>(cancellation.Token));
    }

    [Fact]
    public async Task RowsAsAsync_Explicit_DoesNotYieldMappedRowAfterCancellation()
    {
        using var cancellation = new CancellationTokenSource();
        using var reader = CreateCancelingReader(cancellation);

        await AssertCanceledBeforeYield(reader.RowsAsAsync<AsyncProbeRow>(
            map => map.FromColumn<int>("Id", (row, value) => {
                row.Id = value;
                return row;
            }),
            cancellation.Token));
    }

    [Fact]
    public async Task RowsAsAsync_Factory_DoesNotYieldMappedRowAfterCancellation()
    {
        using var cancellation = new CancellationTokenSource();
        using var reader = CreateCancelingReader(cancellation);

        await AssertCanceledBeforeYield(reader.RowsAsAsync(
            row => row.GetInt32(0),
            cancellation.Token));
    }

    private static ThrowingGetValuesDataReader CreateCancelingReader(
        CancellationTokenSource cancellation) =>
        new(
            ["Id"],
            [[1]],
            afterValueRead: _ => cancellation.Cancel());

    private static async Task AssertCanceledBeforeYield<T>(IAsyncEnumerable<T> rows)
    {
        await using IAsyncEnumerator<T> enumerator = rows.GetAsyncEnumerator();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => enumerator.MoveNextAsync().AsTask());
    }

    private sealed class AsyncProbeRow
    {
        public int Id { get; set; }
    }
}
#endif
