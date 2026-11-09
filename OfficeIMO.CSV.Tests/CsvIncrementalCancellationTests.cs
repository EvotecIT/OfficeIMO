#if NET8_0_OR_GREATER
using System.Data.Common;
using System.Text;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public sealed class CsvIncrementalCancellationTests {
    [Theory]
    [InlineData(false, false, 0)]
    [InlineData(false, true, 0)]
    [InlineData(true, false, 0)]
    [InlineData(true, true, 0)]
    [InlineData(false, false, 1)]
    [InlineData(false, true, 1)]
    [InlineData(false, false, 2)]
    public async Task CancelledAdvanceHidesCurrentRowAndEndsReader(bool parallel, bool lifetimeCancellation, int cancellationStage) {
        using var input = new MemoryStream(Encoding.UTF8.GetBytes("Id,Name\n1,Alpha\n2,Beta\n"));
        using var lifetime = new CancellationTokenSource();
        using var operation = new CancellationTokenSource();
        var callbackCancellation = new OperationCanceledException("The progress callback canceled the read.");
        using DbDataReader reader = await CsvDocument.OpenReaderAsync(input,
            options: new CsvReaderOptions {
                Schema = CreateSchema(),
                ProgressReportInterval = 1,
                ProgressCallback = progress => {
                    if (progress.RecordsRead != 3 || cancellationStage == 0) return;
                    if (cancellationStage == 2) throw callbackCancellation;
                    if (lifetimeCancellation) lifetime.Cancel(); else operation.Cancel();
                },
                ParallelProcessing = parallel ? new CsvDataReaderParallelOptions {
                    BatchSize = 1, MaxDegreeOfParallelism = 2
                } : null
            },
            cancellationToken: lifetime.Token);
        Assert.True(await reader.ReadAsync());
        Assert.Equal(1, reader.GetInt32(0));
        Assert.Equal("Alpha", reader.GetString(1));
        if (cancellationStage == 0) {
            if (lifetimeCancellation) lifetime.Cancel(); else operation.Cancel();
        }
        Task<bool> read = reader.ReadAsync(operation.Token);
        var canceled = await Assert.ThrowsAnyAsync<OperationCanceledException>(() => read);
        Assert.True(read.IsCanceled);
        if (cancellationStage == 2) Assert.Same(callbackCancellation, canceled);
        Assert.Throws<InvalidOperationException>(() => reader.GetString(0));
        await Assert.ThrowsAsync<InvalidOperationException>(() => reader.ReadAsync());
        Assert.True(input.CanRead);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SynchronousAdvanceAfterLifetimeCancellationEndsReader(bool parallel) {
        using var input = new MemoryStream(Encoding.UTF8.GetBytes("Id,Name\n1,Alpha\n2,Beta\n"));
        using var lifetime = new CancellationTokenSource();
        using DbDataReader reader = await CsvDocument.OpenReaderAsync(input,
            options: new CsvReaderOptions {
                Schema = CreateSchema(),
                ParallelProcessing = parallel ? new CsvDataReaderParallelOptions {
                    BatchSize = 1, MaxDegreeOfParallelism = 2
                } : null
            },
            cancellationToken: lifetime.Token);
        Assert.True(reader.Read());
        lifetime.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => reader.Read());
        Assert.Throws<InvalidOperationException>(() => reader.GetString(0));
        Assert.Throws<InvalidOperationException>(() => reader.Read());
    }

    private static CsvSchema CreateSchema() => new CsvSchemaBuilder()
        .Column("Id").AsInt32().Column("Name").AsString().Done().Build();
}
#endif
