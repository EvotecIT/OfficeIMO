#if NET8_0_OR_GREATER
using System;
using System.Data;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Apache.Arrow;
using Apache.Arrow.Ipc;
using OfficeIMO.Data.Arrow;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public sealed class ArrowColumnNullabilityTests {
    [Fact]
    public void OmittedNullabilityKeepsNullableFieldsAndNullValues() {
        using DataTable table = CreateTable();
        table.Rows.Add(1, DBNull.Value);
        using DbDataReader reader = table.CreateDataReader();
        using RecordBatch batch = Assert.Single(reader.ReadArrowBatches());

        Assert.All(batch.Schema.FieldsList, field => Assert.True(field.IsNullable));
        Assert.True(batch.Column(1).IsNull(0));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public async Task RequiredAndNullableFieldsPreserveValuesAndSnapshotSettings(int route) {
        using DataTable table = CreateTable();
        table.Rows.Add(1, "Alpha");
        table.Rows.Add(2, DBNull.Value);
        using DbDataReader reader = table.CreateDataReader();
        bool[] nullability = [false, true];
        var options = new ArrowReadOptions { BatchSize = 1, ColumnNullability = nullability };
        int batches = 0;

        int rows = await ConsumeAsync(reader, options, route, batch => {
            Assert.False(batch.Schema.GetFieldByIndex(0).IsNullable);
            Assert.True(batch.Schema.GetFieldByIndex(1).IsNullable);
            Assert.Equal(++batches, Assert.IsType<Int32Array>(batch.Column(0)).GetValue(0));
            if (batches == 1) {
                Assert.Equal("Alpha", Assert.IsType<StringArray>(batch.Column(1)).GetString(0));
                nullability[1] = false;
                options.ColumnNullability = [true, false];
            } else {
                Assert.True(batch.Column(1).IsNull(0));
            }
        });

        Assert.Equal(2, rows);
        Assert.Equal(2, batches);
        Assert.False(reader.IsClosed);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public async Task RequiredNullIsRejectedBeforePublishingItsBatchDespiteOptionsMutation(int route) {
        using DataTable table = CreateTable();
        table.Rows.Add(1, "Alpha");
        table.Rows.Add(DBNull.Value, "Beta");
        using DbDataReader reader = table.CreateDataReader();
        bool[] nullability = [false, true];
        var options = new ArrowReadOptions { BatchSize = 1, ColumnNullability = nullability };
        int publishedBatches = 0;

        Exception exception = await Assert.ThrowsAnyAsync<Exception>(() =>
            ConsumeAsync(reader, options, route, batch => {
                publishedBatches++;
                Assert.False(batch.Schema.GetFieldByIndex(0).IsNullable);
                Assert.Equal(1, Assert.IsType<Int32Array>(batch.Column(0)).GetValue(0));
                nullability[0] = true;
                options.ColumnNullability = null;
            }));

        if (route != 3) Assert.IsType<InvalidDataException>(exception);
        Assert.Contains("Required Arrow column 'Id'", exception.Message, StringComparison.Ordinal);
        Assert.Equal(1, publishedBatches);
        Assert.False(reader.IsClosed);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public async Task NullabilityWidthIsRejectedBeforeSourceRowsAreRead(int route) {
        using DataTable table = CreateTable();
        table.Rows.Add(1, "Alpha");
        using DbDataReader reader = table.CreateDataReader();
        var options = new ArrowReadOptions { ColumnNullability = [false] };

        ArgumentException exception = await Assert.ThrowsAsync<ArgumentException>(() =>
            ConsumeAsync(reader, options, route, _ => throw new InvalidOperationException("Unexpected batch.")));

        Assert.Equal(nameof(ArrowReadOptions.ColumnNullability), exception.ParamName);
        Assert.True(reader.Read());
        Assert.Equal(1, reader.GetInt32(0));
    }

    private static DataTable CreateTable() {
        var table = new DataTable();
        table.Columns.Add("Id", typeof(int));
        table.Columns.Add("Name", typeof(string));
        return table;
    }

    // All routes use the public adapter and retain caller ownership of the source reader.
    private static async Task<int> ConsumeAsync(
        DbDataReader reader, ArrowReadOptions? options, int route, Action<RecordBatch> consume) {
        int rows = 0;
        void Observe(RecordBatch batch) {
            using (batch) {
                consume(batch);
                rows += batch.Length;
            }
        }

        if (route == 0) {
            foreach (RecordBatch batch in reader.ReadArrowBatches(options)) Observe(batch);
        } else if (route == 1) {
            await foreach (RecordBatch batch in reader.ReadArrowBatchesAsync(options)) Observe(batch);
        } else {
            using ArrowCArrayStreamOwner? owner = route == 3 ? reader.ExportArrowCStream(options) : null;
            using IArrowArrayStream stream = owner is null ? reader.OpenArrowStream(options) : owner.ImportArrayStream();
            while (await stream.ReadNextRecordBatchAsync() is { } batch) Observe(batch);
        }
        return rows;
    }
}
#endif
