using System;
using System.Data.Common;
using System.Data;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public sealed class CsvParallelDataReaderTests
{
    [Fact]
    public void OpenTextDataReader_ParallelTypedProjectionPreservesOrderSchemaAndHasRows()
    {
        const int rowCount = 5000;
        var csv = new StringBuilder("Id,Amount,Active,Created,Name\n");
        for (int id = 1; id <= rowCount; id++)
        {
            csv.Append(id).Append(',')
                .Append((id * 1.25m).ToString(CultureInfo.InvariantCulture)).Append(',')
                .Append((id & 1) == 0 ? "true" : "false").Append(',')
                .Append("2026-08-").Append(((id % 28) + 1).ToString("00", CultureInfo.InvariantCulture)).Append(',')
                .Append("Row ").Append(id).Append('\n');
        }

        CsvSchema schema = CreateTypedSchema();
        using DbDataReader reader = CsvDocument.OpenTextDataReader(
            csv.ToString(),
            readerOptions: new CsvDataReaderOptions
            {
                Schema = schema,
                ParallelProcessing = new CsvDataReaderParallelOptions
                {
                    MaxDegreeOfParallelism = 4,
                    BatchSize = 127
                }
            });

        ICsvDataReaderPositionMetadata position =
            Assert.IsAssignableFrom<ICsvDataReaderPositionMetadata>(reader);
        Assert.True(reader.HasRows);
        Assert.Equal(0, position.RecordNumber);
        Assert.Equal(typeof(int), reader.GetFieldType(0));
        Assert.Equal(typeof(decimal), reader.GetFieldType(1));

        var values = new object[reader.FieldCount];
        int expectedId = 0;
        decimal amountChecksum = 0;
        while (reader.Read())
        {
            expectedId++;
            Assert.Equal(expectedId, reader.GetInt32(0));
            Assert.Equal(expectedId, position.RecordNumber);
            Assert.Equal((expectedId & 1) == 0, reader.GetBoolean(2));
            Assert.Equal($"Row {expectedId}", reader.GetString(4));
            Assert.Equal(reader.FieldCount, reader.GetValues(values));
            amountChecksum += (decimal)values[1];
        }

        Assert.Equal(rowCount, expectedId);
        Assert.Equal(1.25m * rowCount * (rowCount + 1) / 2, amountChecksum);
        Assert.Equal(0, position.RecordNumber);
        Assert.True(reader.HasRows);
    }

    [Fact]
    public void OpenDataReader_ParallelStreamingFallbackPreservesMultilineValuesAndOrder()
    {
        CsvSchema schema = new CsvSchemaBuilder()
            .Column("Id").AsInt32()
            .Column("Value").AsString()
            .Done()
            .Build();
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(
            "Id,Value\n1,one\n2,\"line one\nline two\"\n3,three\n"));
        using DbDataReader reader = CsvDocument.OpenDataReader(
            stream,
            readerOptions: new CsvDataReaderOptions
            {
                Schema = schema,
                ParallelProcessing = new CsvDataReaderParallelOptions
                {
                    MaxDegreeOfParallelism = 2,
                    BatchSize = 2
                }
            });
        ICsvDataReaderPositionMetadata position =
            Assert.IsAssignableFrom<ICsvDataReaderPositionMetadata>(reader);

        Assert.True(reader.Read());
        Assert.Equal(1, reader.GetInt32(0));
        Assert.Equal(1, position.RecordNumber);

        Assert.True(reader.Read());
        Assert.Equal("line one\nline two", reader.GetString(1));
        Assert.Equal(2, position.RecordNumber);

        Assert.True(reader.Read());
        Assert.Equal(3, reader.GetInt32(0));
        Assert.Equal(3, position.RecordNumber);
        Assert.False(reader.Read());
        Assert.True(stream.CanRead);
    }

    [Fact]
    public void OpenTextDataReader_ParallelConversionYieldsValidPrefixBeforeOrderedError()
    {
        CsvSchema schema = new CsvSchemaBuilder()
            .Column("Id").AsInt32()
            .Done()
            .Build();
        using DbDataReader reader = CsvDocument.OpenTextDataReader(
            "Id\n1\n2\nnot-an-integer\n4\n",
            readerOptions: new CsvDataReaderOptions
            {
                Schema = schema,
                ParallelProcessing = new CsvDataReaderParallelOptions
                {
                    MaxDegreeOfParallelism = 3,
                    BatchSize = 2
                }
            });

        Assert.True(reader.Read());
        Assert.Equal(1, reader.GetInt32(0));
        Assert.True(reader.Read());
        Assert.Equal(2, reader.GetInt32(0));
        CsvException exception = Assert.Throws<CsvException>(() => reader.Read());
        Assert.Contains("not-an-integer", exception.Message);
        Assert.Contains("row 3", exception.Message);
    }

    [Fact]
    public async Task OpenTextDataReader_ParallelReadAsyncObservesCallerCancellation()
    {
        CsvSchema schema = new CsvSchemaBuilder()
            .Column("Id").AsInt32()
            .Done()
            .Build();
        using DbDataReader reader = CsvDocument.OpenTextDataReader(
            "Id\n1\n2\n",
            readerOptions: new CsvDataReaderOptions
            {
                Schema = schema,
                ParallelProcessing = new CsvDataReaderParallelOptions
                {
                    MaxDegreeOfParallelism = 2,
                    BatchSize = 1
                }
            });
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => reader.ReadAsync(cancellation.Token));
    }

    [Fact]
    public void OpenTextDataReader_ParallelReaderObservesLoadCancellationBetweenBufferedRows()
    {
        CsvSchema schema = new CsvSchemaBuilder()
            .Column("Id").AsInt32()
            .Done()
            .Build();
        using var cancellation = new CancellationTokenSource();
        using DbDataReader reader = CsvDocument.OpenTextDataReader(
            "Id\n1\n2\n3\n",
            new CsvLoadOptions { CancellationToken = cancellation.Token },
            new CsvDataReaderOptions
            {
                Schema = schema,
                ParallelProcessing = new CsvDataReaderParallelOptions
                {
                    MaxDegreeOfParallelism = 2,
                    BatchSize = 2
                }
            });

        Assert.True(reader.Read());
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => reader.Read());
    }

    [Fact]
    public void OpenDataReader_ParallelInferredFallbackRetainsLoadCancellation()
    {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.CSV.ParallelCancel.{Guid.NewGuid():N}.csv.gz");
        using var cancellation = new CancellationTokenSource();

        try
        {
            WriteGZipCsv(path, "Id\n1\n2\n3\n4\n5\n6\n");
            using DbDataReader reader = CsvDocument.OpenDataReader(
                path,
                new CsvLoadOptions
                {
                    CompressionType = CsvCompressionType.GZip,
                    CancellationToken = cancellation.Token
                },
                new CsvDataReaderOptions
                {
                    InferSchema = true,
                    SchemaSampleSize = 1,
                    ParallelProcessing = new CsvDataReaderParallelOptions
                    {
                        MaxDegreeOfParallelism = 2,
                        BatchSize = 2
                    }
                });

            Assert.True(reader.Read());
            Assert.Equal(1, reader.GetInt32(0));
            cancellation.Cancel();
            Assert.Throws<OperationCanceledException>(() => reader.Read());
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void OpenTextDataReader_ParallelReaderKeepsProducerPrefetchBounded()
    {
        var csv = new StringBuilder("Id\n");
        for (int id = 1; id <= 1000; id++) csv.Append(id).Append('\n');
        int recordsReported = 0;
        CsvSchema schema = new CsvSchemaBuilder()
            .Column("Id").AsInt32()
            .Done()
            .Build();
        using DbDataReader reader = CsvDocument.OpenTextDataReader(
            csv.ToString(),
            new CsvLoadOptions
            {
                ProgressReportInterval = 1,
                ProgressCallback = progress => recordsReported = checked((int)progress.RecordsRead)
            },
            new CsvDataReaderOptions
            {
                Schema = schema,
                ParallelProcessing = new CsvDataReaderParallelOptions
                {
                    MaxDegreeOfParallelism = 2,
                    BatchSize = 8
                }
            });

        Assert.True(reader.Read());
        Assert.InRange(recordsReported, 1, 16);
        Assert.Equal(1, reader.GetInt32(0));
    }

    [Fact]
    public void OpenTextDataReader_ParallelReaderLoadsTheSameTypedDataTable()
    {
        CsvSchema schema = CreateTypedSchema();
        const string csv =
            "Id,Amount,Active,Created,Name\n" +
            "1,1.25,false,2026-08-02,Row 1\n" +
            "2,2.50,true,2026-08-03,Row 2\n";
        using DbDataReader reader = CsvDocument.OpenTextDataReader(
            csv,
            readerOptions: new CsvDataReaderOptions
            {
                Schema = schema,
                ParallelProcessing = new CsvDataReaderParallelOptions
                {
                    MaxDegreeOfParallelism = 2,
                    BatchSize = 1
                }
            });
        var table = new DataTable();

        table.Load(reader);

        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(typeof(int), table.Columns["Id"]!.DataType);
        Assert.Equal(typeof(decimal), table.Columns["Amount"]!.DataType);
        Assert.Equal(2, table.Rows[1]["Id"]);
        Assert.Equal(2.50m, table.Rows[1]["Amount"]);
    }

    [Fact]
    public void OpenDataReader_ParallelReaderDisposalReleasesTheSourceFile()
    {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.CSV.ParallelDispose.{Guid.NewGuid():N}.csv.gz");
        WriteGZipCsv(path, "Id\n1\n2\n3\n");
        CsvSchema schema = new CsvSchemaBuilder()
            .Column("Id").AsInt32()
            .Done()
            .Build();

        try
        {
            using (DbDataReader reader = CsvDocument.OpenDataReader(
                       path,
                       new CsvLoadOptions { CompressionType = CsvCompressionType.GZip },
                       readerOptions: new CsvDataReaderOptions
                       {
                           Schema = schema,
                           ParallelProcessing = new CsvDataReaderParallelOptions
                           {
                               MaxDegreeOfParallelism = 2,
                               BatchSize = 1
                           }
                       }))
            {
                Assert.True(reader.Read());
            }

            File.Delete(path);
            Assert.False(File.Exists(path));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Theory]
    [InlineData(0, 16)]
    [InlineData(2, 0)]
    public void OpenTextDataReader_RejectsInvalidParallelSettings(int degree, int batchSize)
    {
        CsvSchema schema = new CsvSchemaBuilder()
            .Column("Id").AsInt32()
            .Done()
            .Build();

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            CsvDocument.OpenTextDataReader(
                "Id\n1\n",
                readerOptions: new CsvDataReaderOptions
                {
                    Schema = schema,
                    ParallelProcessing = new CsvDataReaderParallelOptions
                    {
                        MaxDegreeOfParallelism = degree,
                        BatchSize = batchSize
                    }
                }));
    }

    private static CsvSchema CreateTypedSchema() => new CsvSchemaBuilder()
        .Column("Id").AsInt32()
        .Column("Amount").AsType(typeof(decimal))
        .Column("Active").AsBoolean()
        .Column("Created").AsDateTime()
        .Column("Name").AsString()
        .Done()
        .Build();

    private static void WriteGZipCsv(string path, string text)
    {
        using var file = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        using var gzip = new GZipStream(file, CompressionLevel.Fastest);
        using var writer = new StreamWriter(gzip, new UTF8Encoding(false));
        writer.Write(text);
    }
}
