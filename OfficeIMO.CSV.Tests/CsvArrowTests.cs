#if NET8_0_OR_GREATER
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Apache.Arrow;
using Apache.Arrow.C;
using Apache.Arrow.Ipc;
using Apache.Arrow.Types;
using OfficeIMO.Data.Arrow;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public sealed class CsvArrowTests {
    [Fact]
    public async Task ManagedArrowStreamIsBoundedSequentialAndLeavesReaderOpen() {
        CsvDocument document = CsvDocument.Parse(
            "Id,Name\n" +
            "1,Alpha\n" +
            "2,Beta\n" +
            "3,Gamma\n");
        using var reader = document.CreateDataReader(new CsvDataReaderOptions { InferSchema = true });
        using IArrowArrayStream stream = reader.OpenArrowStream(
            new ArrowReadOptions { BatchSize = 2 });

        Assert.Equal(2, stream.Schema.FieldsList.Count);
        using RecordBatch first = (await stream.ReadNextRecordBatchAsync())!;
        using RecordBatch second = (await stream.ReadNextRecordBatchAsync())!;

        Assert.Equal(2, first.Length);
        Assert.Equal(1, second.Length);
        Assert.Null(await stream.ReadNextRecordBatchAsync());
        Assert.False(reader.IsClosed);
    }

    [Fact]
    public async Task ArrowCStreamRoundTripsBoundedBatchesAndMovesOwnershipSafely() {
        CsvDocument document = CsvDocument.Parse(
            "Id,Name\n" +
            "1,Alpha\n" +
            "2,Beta\n" +
            "3,Gamma\n");
        using var reader = document.CreateDataReader(new CsvDataReaderOptions { InferSchema = true });
        using ArrowCArrayStreamOwner owner = reader.ExportArrowCStream(
            new ArrowReadOptions { BatchSize = 2 });

        Assert.NotEqual(0, owner.Address);
        IArrowArrayStream importedStream;
        unsafe {
            importedStream = CArrowArrayStreamImporter.ImportArrayStream(
                owner.DangerousGetPointer());
        }
        using IArrowArrayStream imported = importedStream;

        // Import moves the release callback into the managed importer. The owner can
        // now free its original struct without releasing the moved stream twice.
        owner.Dispose();
        Assert.True(owner.IsDisposed);
        Assert.Equal(2, imported.Schema.FieldsList.Count);

        using RecordBatch first = (await imported.ReadNextRecordBatchAsync())!;
        using RecordBatch second = (await imported.ReadNextRecordBatchAsync())!;
        Assert.Equal(2, first.Length);
        Assert.Equal(1, second.Length);
        Assert.Equal(1, Assert.IsType<Int32Array>(first.Column(0)).GetValue(0));
        Assert.Equal("Gamma", Assert.IsType<StringArray>(second.Column(1)).GetString(0));
        Assert.Null(await imported.ReadNextRecordBatchAsync());
        Assert.False(reader.IsClosed);
    }

    [Fact]
    public async Task ManagedArrowStreamRejectsReadsAfterDisposal() {
        CsvDocument document = CsvDocument.Parse("Id\n1\n");
        using var reader = document.CreateDataReader(new CsvDataReaderOptions { InferSchema = true });
        IArrowArrayStream stream = reader.OpenArrowStream();
        stream.Dispose();

        await Assert.ThrowsAsync<ObjectDisposedException>(() =>
            stream.ReadNextRecordBatchAsync().AsTask());
        Assert.False(reader.IsClosed);
    }

    [Fact]
    public async Task InferredCsvReaderStreamsTypedBoundedArrowBatches() {
        CsvDocument document = CsvDocument.Parse(
            "Id,Amount,Active,Created,Name\n" +
            "1,12.50,true,2026-08-30,Alpha\n" +
            "2,,false,2026-08-31,Beta\n" +
            "3,99.25,true,2026-09-01,Gamma\n");
        using var reader = document.CreateDataReader(new CsvDataReaderOptions { InferSchema = true });
        var batches = new List<RecordBatch>();

        await foreach (RecordBatch batch in reader.ReadArrowBatchesAsync(
                           new ArrowReadOptions { BatchSize = 2, DecimalScale = 2 })) {
            batches.Add(batch);
        }

        try {
            Assert.Equal(2, batches.Count);
            Assert.Equal(2, batches[0].Length);
            Assert.Equal(1, batches[1].Length);
            Assert.IsType<Int32Type>(batches[0].Schema.GetFieldByIndex(0).DataType);
            Assert.IsType<Decimal128Type>(batches[0].Schema.GetFieldByIndex(1).DataType);
            Assert.IsType<BooleanType>(batches[0].Schema.GetFieldByIndex(2).DataType);
            Assert.IsType<TimestampType>(batches[0].Schema.GetFieldByIndex(3).DataType);
            Assert.IsType<StringType>(batches[0].Schema.GetFieldByIndex(4).DataType);

            var ids = Assert.IsType<Int32Array>(batches[0].Column(0));
            var names = Assert.IsType<StringArray>(batches[1].Column(4));
            Assert.Equal(1, ids.GetValue(0));
            Assert.Equal(2, ids.GetValue(1));
            Assert.Equal("Gamma", names.GetString(0));
            Assert.True(batches[0].Column(1).IsNull(1));
        } finally {
            foreach (RecordBatch batch in batches) batch.Dispose();
        }
    }

    [Fact]
    public void ArrowAdapterCanFailClosedForUnsupportedClrTypes() {
        var table = new System.Data.DataTable();
        table.Columns.Add("Value", typeof(Uri));
        table.Rows.Add(new Uri("https://example.test/"));
        using var reader = table.CreateDataReader();

        Assert.Throws<NotSupportedException>(() => reader.ReadArrowBatches(
            new ArrowReadOptions { ConvertUnsupportedTypesToString = false }).ToArray());
    }

    [Fact]
    public void ArrowDecimalOptionsAreInitializerOrderIndependentAndValidateBeforeReading() {
        var valid = new ArrowReadOptions { DecimalPrecision = 5, DecimalScale = 2 };
        Assert.Equal(5, valid.DecimalPrecision);
        Assert.Equal(2, valid.DecimalScale);

        var invalid = new ArrowReadOptions { DecimalPrecision = 10, DecimalScale = 12 };
        var table = new System.Data.DataTable();
        table.Columns.Add("Value", typeof(decimal));
        table.Rows.Add(1m);
        using var reader = table.CreateDataReader();

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            reader.ReadArrowBatches(invalid).ToArray());
    }

    [Fact]
    public void ExplicitColumnTypesSkipReaderSchemaInferenceAndAreValidated() {
        var table = new System.Data.DataTable();
        table.Columns.Add("Id", typeof(object));
        table.Rows.Add(42);
        using var reader = table.CreateDataReader();

        RecordBatch batch = Assert.Single(reader.ReadArrowBatches(
            new ArrowReadOptions { ColumnTypes = new[] { typeof(int) } }));
        try {
            Assert.IsType<Int32Type>(batch.Schema.GetFieldByIndex(0).DataType);
            Assert.Equal(42, Assert.IsType<Int32Array>(batch.Column(0)).GetValue(0));
        } finally {
            batch.Dispose();
        }

        using var invalidReader = table.CreateDataReader();
        Assert.Throws<ArgumentException>(() => invalidReader.ReadArrowBatches(
            new ArrowReadOptions { ColumnTypes = new[] { typeof(int), typeof(string) } }).ToArray());
    }

    [Fact]
    public void WideArrowReaderBoundsInitialReservationAcrossColumns() {
        const int columnCount = 1024;
        var table = new System.Data.DataTable();
        object[] values = new object[columnCount];
        for (int ordinal = 0; ordinal < columnCount; ordinal++) {
            table.Columns.Add("Column" + ordinal, typeof(int));
            values[ordinal] = ordinal;
        }
        table.Rows.Add(values);
        using var reader = table.CreateDataReader();

        RecordBatch batch = Assert.Single(reader.ReadArrowBatches());
        try {
            Assert.Equal(1, batch.Length);
            Assert.Equal(columnCount, batch.ColumnCount);
        } finally {
            batch.Dispose();
        }
    }

    [Fact]
    public void ArrowAdapterSeparatesTimezoneLessDateTimeFromUtcDateTimeOffset() {
        DateTime wallClock = new(2026, 8, 31, 14, 35, 12, DateTimeKind.Unspecified);
        DateTimeOffset instant = new(2026, 8, 31, 14, 35, 12, TimeSpan.FromHours(2));
        var table = new System.Data.DataTable();
        table.Columns.Add("WallClock", typeof(DateTime));
        table.Columns.Add("Instant", typeof(DateTimeOffset));
        table.Rows.Add(wallClock, instant);
        using var reader = table.CreateDataReader();

        RecordBatch batch = Assert.Single(reader.ReadArrowBatches());
        try {
            var wallClockType = Assert.IsType<TimestampType>(batch.Schema.GetFieldByIndex(0).DataType);
            var instantType = Assert.IsType<TimestampType>(batch.Schema.GetFieldByIndex(1).DataType);
            Assert.False(wallClockType.IsTimeZoneAware);
            Assert.True(instantType.IsTimeZoneAware);

            var wallClockValues = Assert.IsType<TimestampArray>(batch.Column(0));
            var instantValues = Assert.IsType<TimestampArray>(batch.Column(1));
            Assert.Equal(wallClock.Ticks, wallClockValues.GetTimestamp(0)!.Value.UtcTicks);
            Assert.Equal(instant.UtcTicks, instantValues.GetTimestamp(0)!.Value.UtcTicks);
        } finally {
            batch.Dispose();
        }
    }

    [Fact]
    public void SyncArrowAdapterRejectsPreCancelledEmptyReader() {
        var table = new System.Data.DataTable();
        table.Columns.Add("Value", typeof(int));
        using var reader = table.CreateDataReader();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            reader.ReadArrowBatches(cancellationToken: cancellation.Token).ToArray());
    }

    [Fact]
    public async Task AsyncArrowAdapterRejectsCancellationRaisedBySuccessfulRead() {
        using var cancellation = new CancellationTokenSource();
        using var reader = new ThrowingGetValuesDataReader(
            new[] { "Value" },
            new[] { new object?[] { 42 } },
            afterRead: _ => cancellation.Cancel());
        await using IAsyncEnumerator<RecordBatch> batches = reader
            .ReadArrowBatchesAsync(
                new ArrowReadOptions { BatchSize = 1 },
                cancellation.Token)
            .GetAsyncEnumerator();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            batches.MoveNextAsync().AsTask());
    }

    [Fact]
    public void SyncArrowAdapterRejectsCancellationRaisedByEndOfDataRead() {
        using var cancellation = new CancellationTokenSource();
        using var reader = new ThrowingGetValuesDataReader(
            new[] { "Value" },
            new[] { new object?[] { 42 } },
            afterEnd: cancellation.Cancel);

        Assert.Throws<OperationCanceledException>(() =>
            reader.ReadArrowBatches(
                new ArrowReadOptions { BatchSize = 2 },
                cancellation.Token)
                .ToArray());
    }

    [Fact]
    public void SyncArrowAdapterStopsBeforeNextReadWhenGetterCancels() {
        using var cancellation = new CancellationTokenSource();
        int readCount = 0;
        using var reader = new ThrowingGetValuesDataReader(
            new[] { "Value" },
            new[] { new object?[] { 42 }, new object?[] { 43 } },
            afterRead: _ => readCount++,
            afterValueRead: _ => cancellation.Cancel());

        Assert.Throws<OperationCanceledException>(() =>
            reader.ReadArrowBatches(
                new ArrowReadOptions { BatchSize = 2 },
                cancellation.Token)
                .ToArray());
        Assert.Equal(1, readCount);
    }
}
#endif
