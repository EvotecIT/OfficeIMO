#if NET8_0_OR_GREATER
using System;
using System.IO;
using System.Data.Common;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public sealed class CsvIncrementalCancellationTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task CancelledAdvanceHidesCurrentRowAndEndsReader(bool parallel, bool lifetimeCancellation) {
        using var input = new MemoryStream(Encoding.UTF8.GetBytes("Name\nAlpha\nBeta\n"));
        using var lifetime = new CancellationTokenSource();
        using var operation = new CancellationTokenSource();
        using DbDataReader reader = await CsvDocument.OpenDataReaderAsync(input,
            cancellationToken: lifetime.Token,
            readerOptions: new CsvDataReaderOptions {
                ParallelProcessing = parallel ? new CsvDataReaderParallelOptions {
                    BatchSize = 1, MaxDegreeOfParallelism = 2
                } : null
            });
        Assert.True(await reader.ReadAsync());
        Assert.Equal("Alpha", reader.GetString(0));
        if (lifetimeCancellation) lifetime.Cancel(); else operation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => reader.ReadAsync(operation.Token));
        Assert.Throws<InvalidOperationException>(() => reader.GetString(0));
        await Assert.ThrowsAsync<InvalidOperationException>(() => reader.ReadAsync());
        Assert.True(input.CanRead);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SynchronousAdvanceAfterLifetimeCancellationEndsReader(bool parallel) {
        using var input = new MemoryStream(Encoding.UTF8.GetBytes("Name\nAlpha\nBeta\n"));
        using var lifetime = new CancellationTokenSource();
        using DbDataReader reader = await CsvDocument.OpenDataReaderAsync(input,
            cancellationToken: lifetime.Token,
            readerOptions: new CsvDataReaderOptions {
                ParallelProcessing = parallel ? new CsvDataReaderParallelOptions {
                    BatchSize = 1, MaxDegreeOfParallelism = 2
                } : null
            });
        Assert.True(reader.Read());
        lifetime.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => reader.Read());
        Assert.Throws<InvalidOperationException>(() => reader.GetString(0));
        Assert.Throws<InvalidOperationException>(() => reader.Read());
    }
}
#endif
