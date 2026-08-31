#if NET8_0_OR_GREATER
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Apache.Arrow;
using Apache.Arrow.Types;
using OfficeIMO.Data.Arrow;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public sealed class CsvArrowTests {
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
    public void ArrowDecimalOptionsCannotEnterAnInvalidPrecisionScaleState() {
        var options = new ArrowReadOptions { DecimalScale = 12 };

        Assert.Throws<ArgumentOutOfRangeException>(() => options.DecimalPrecision = 10);
        Assert.Equal(29, options.DecimalPrecision);
        Assert.Equal(12, options.DecimalScale);
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
}
#endif
