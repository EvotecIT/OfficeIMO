using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvDataReaderTests
{
    [Fact]
    public void CreateDataReader_WithInferredSchema_ExposesTypedColumnsAndValues()
    {
        var doc = CsvDocument.Parse("Id,Amount,Active,Created,Note\n1,12.5,true,2026-07-07,Alpha\n2,,false,2026-07-08,\n");

        using var reader = doc.CreateDataReader(new CsvDataReaderOptions { InferSchema = true });

        Assert.True(reader.HasRows);
        Assert.Equal(5, reader.FieldCount);
        Assert.Equal(typeof(int), reader.GetFieldType(0));
        Assert.Equal(typeof(decimal), reader.GetFieldType(1));
        Assert.Equal(typeof(bool), reader.GetFieldType(2));
        Assert.Equal(typeof(DateTime), reader.GetFieldType(3));
        Assert.Equal(typeof(string), reader.GetFieldType(4));

        Assert.True(reader.Read());
        Assert.Equal(1, reader.GetInt32(reader.GetOrdinal("Id")));
        Assert.Equal(12.5m, reader.GetDecimal(reader.GetOrdinal("Amount")));
        Assert.True(reader.GetBoolean(reader.GetOrdinal("Active")));

        Assert.True(reader.Read());
        Assert.True(reader.IsDBNull(reader.GetOrdinal("Amount")));
        Assert.Equal(string.Empty, reader.GetString(reader.GetOrdinal("Note")));
        Assert.False(reader.Read());
    }

    [Fact]
    public void DataTableLoad_FromCsvDataReader_UsesTypedSchema()
    {
        var doc = CsvDocument.Parse("Id,Amount,Active\n1,12.5,true\n2,13.75,false\n");

        using var reader = doc.CreateDataReader(new CsvDataReaderOptions { InferSchema = true });
        var table = new DataTable();
        table.Load(reader);

        Assert.Equal(typeof(int), table.Columns["Id"]!.DataType);
        Assert.Equal(typeof(decimal), table.Columns["Amount"]!.DataType);
        Assert.Equal(typeof(bool), table.Columns["Active"]!.DataType);
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(1, table.Rows[0]["Id"]);
        Assert.Equal(13.75m, table.Rows[1]["Amount"]);
    }

    [Fact]
    public void CreateDataReader_WithExplicitSchemaBuilder_UsesTypedColumnsWithoutInference()
    {
        var schema = new CsvSchemaBuilder()
            .Column("Id").AsInt32()
            .Column("Amount").AsType(typeof(decimal))
            .Column("Created").AsDateTime()
            .Done()
            .Build();
        var doc = CsvDocument.Parse("Id,Amount,Created\n1,12.5,2026-07-07\n");

        using var reader = doc.CreateDataReader(new CsvDataReaderOptions { Schema = schema });

        Assert.Equal(typeof(int), reader.GetFieldType(reader.GetOrdinal("Id")));
        Assert.Equal(typeof(decimal), reader.GetFieldType(reader.GetOrdinal("Amount")));
        Assert.Equal(typeof(DateTime), reader.GetFieldType(reader.GetOrdinal("Created")));
        Assert.True(reader.Read());
        Assert.Equal(1, reader.GetInt32(reader.GetOrdinal("Id")));
        Assert.Equal(12.5m, reader.GetDecimal(reader.GetOrdinal("Amount")));
        Assert.Equal(new DateTime(2026, 7, 7), reader.GetDateTime(reader.GetOrdinal("Created")));
    }

    [Fact]
    public void TypedGetters_RequireAnOpenCurrentRow()
    {
        var schema = new CsvSchemaBuilder()
            .Column("Id").AsInt32()
            .Done()
            .Build();
        var doc = CsvDocument.Parse(
            "Id\n1\n",
            new CsvLoadOptions { Mode = CsvLoadMode.Stream });
        using var reader = doc.CreateDataReader(new CsvDataReaderOptions { Schema = schema });

        Assert.Throws<InvalidOperationException>(() => reader.GetInt32(0));
        Assert.True(reader.Read());
        Assert.Equal(1, reader.GetInt32(0));

        reader.Close();

        Assert.Throws<InvalidOperationException>(() => reader.GetInt32(0));
    }

    [Fact]
    public void CreateDataReader_FromFile_DisposeWithoutRead_ReleasesFile()
    {
        var path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.Tests." + Guid.NewGuid().ToString("N") + ".csv");
        File.WriteAllText(path, "Id,Name\n1,Alice\n2,Bob\n");

        try
        {
            using (var reader = CsvDocument.CreateDataReader(path))
            {
                using var schema = reader.GetSchemaTable();
                Assert.Equal(2, reader.FieldCount);
            }

            using var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            Assert.True(stream.CanRead);
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void GetValues_WithInferredSchema_ExposesTypedValues()
    {
        var doc = CsvDocument.Parse("Id,Amount,Created,Note\n1,12.5,2026-07-07,Alpha\n");

        using var reader = doc.CreateDataReader(new CsvDataReaderOptions { InferSchema = true });

        Assert.True(reader.Read());
        var values = new object[reader.FieldCount];
        Assert.Equal(4, reader.GetValues(values));
        Assert.Equal(1, values[0]);
        Assert.Equal(12.5m, values[1]);
        Assert.Equal(new DateTime(2026, 7, 7), values[2]);
        Assert.Equal("Alpha", values[3]);
    }

    [Fact]
    public void GetValues_WithExplicitSchema_UsesCommonTypedConversions()
    {
        var guid = Guid.Parse("2fae048c-5886-43ec-b03f-e5814c5d52ba");
        var schema = new CsvSchemaBuilder()
            .Column("Id").AsInt32()
            .Column("Name").AsString()
            .Column("Score").AsType(typeof(decimal))
            .Column("Created").AsDateTime()
            .Column("Active").AsBoolean()
            .Column("BatchId").AsType(typeof(Guid))
            .Done()
            .Build();
        var doc = CsvDocument.Parse(
            $"Id,Name,Score,Created,Active,BatchId\n{int.MinValue},Alpha,-12.50,01/02/2026 03:04:05,1,{guid}\n",
            new CsvLoadOptions { Mode = CsvLoadMode.Stream });

        using var reader = doc.CreateDataReader(new CsvDataReaderOptions { Schema = schema });

        Assert.True(reader.Read());
        var values = new object[reader.FieldCount];
        Assert.Equal(6, reader.GetValues(values));
        Assert.Equal(int.MinValue, values[0]);
        Assert.Equal("Alpha", values[1]);
        Assert.Equal(-12.50m, values[2]);
        Assert.Equal(new DateTime(2026, 1, 2, 3, 4, 5), values[3]);
        Assert.Equal(true, values[4]);
        Assert.Equal(guid, values[5]);
    }

    [Fact]
    public void GetValues_WithCustomConverter_CachesConvertedValuesForCurrentRow()
    {
        var calls = 0;
        var doc = CsvDocument.Parse("Score\nhigh\n")
            .EnsureSchema(schema => schema
                .Column("Score")
                .AsInt32()
                .ConvertUsing(value =>
                {
                    calls++;
                    return string.Equals(Convert.ToString(value), "high", StringComparison.OrdinalIgnoreCase) ? 10 : 1;
                }));

        using var reader = doc.CreateDataReader();

        Assert.True(reader.Read());
        var values = new object[reader.FieldCount];
        Assert.Equal(1, reader.GetValues(values));
        Assert.Equal(10, values[0]);
        Assert.Equal(10, reader.GetValue(0));
        Assert.Equal(1, calls);
    }

    [Fact]
    public void DataTableLoad_FromCsvDataReader_HandlesWideSchema()
    {
        var header = string.Join(",", Enumerable.Range(1, 40).Select(i => $"C{i}"));
        var row = string.Join(",", Enumerable.Range(1, 40));
        var doc = CsvDocument.Parse(header + "\n" + row + "\n");

        using var reader = doc.CreateDataReader(new CsvDataReaderOptions { InferSchema = true });
        var table = new DataTable();
        table.Load(reader);

        Assert.Equal(40, table.Columns.Count);
        Assert.Equal(typeof(int), table.Columns["C40"]!.DataType);
        Assert.Equal(40, table.Rows[0]["C40"]);
    }

    [Fact]
    public void CreateDataReader_WithStreamingInferredSchema_ReturnsSampledAndRemainingRows()
    {
        var doc = CsvDocument.Parse(
            "Id,Name\n1,Alpha\n2,Beta\n3,Gamma\n",
            new CsvLoadOptions { Mode = CsvLoadMode.Stream });

        using var reader = doc.CreateDataReader(new CsvDataReaderOptions { InferSchema = true, SchemaSampleSize = 2 });

        Assert.Equal(typeof(int), reader.GetFieldType(reader.GetOrdinal("Id")));
        Assert.True(reader.Read());
        Assert.Equal(1, reader.GetInt32(0));
        Assert.Equal("Alpha", reader.GetString(1));
        Assert.True(reader.Read());
        Assert.Equal(2, reader.GetInt32(0));
        Assert.Equal("Beta", reader.GetString(1));
        Assert.True(reader.Read());
        Assert.Equal(3, reader.GetInt32(0));
        Assert.Equal("Gamma", reader.GetString(1));
        Assert.False(reader.Read());
    }

    [Fact]
    public void CreateDataReader_WithStreamingInferredSchema_EnforcesRequiredColumnsAfterSample()
    {
        var doc = CsvDocument.Parse(
            "Id,Name\n1,Alpha\n2,Beta\n,Gamma\n",
            new CsvLoadOptions { Mode = CsvLoadMode.Stream });

        using var reader = doc.CreateDataReader(new CsvDataReaderOptions { InferSchema = true, SchemaSampleSize = 2 });
        Assert.True(reader.Read());
        Assert.True(reader.Read());
        Assert.True(reader.Read());

        var ex = Assert.Throws<CsvException>(() => reader.GetValue(0));
        Assert.Contains("Column 'Id' is required", ex.Message);
    }

    [Fact]
    public void CreateDataReader_WithLargeStreamingInferenceSample_ReturnsEveryRow()
    {
        const int rowCount = 5000;
        var csv = new StringBuilder("Id,Name\n");
        for (var i = 1; i <= rowCount; i++)
        {
            csv.Append(i).Append(",Name-").Append(i).Append('\n');
        }

        var doc = CsvDocument.Parse(csv.ToString(), new CsvLoadOptions { Mode = CsvLoadMode.Stream });
        using var reader = doc.CreateDataReader(new CsvDataReaderOptions
        {
            InferSchema = true,
            SchemaSampleSize = rowCount
        });

        Assert.Equal(typeof(int), reader.GetFieldType(0));
        var actualRowCount = 0;
        var lastId = 0;
        while (reader.Read())
        {
            actualRowCount++;
            lastId = reader.GetInt32(0);
        }

        Assert.Equal(rowCount, actualRowCount);
        Assert.Equal(rowCount, lastId);
    }

    [Fact]
    public void CreateDataReader_FromFileWithExplicitSchema_ExposesTypedRows()
    {
        var path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.Tests." + Guid.NewGuid().ToString("N") + ".csv");
        var schema = new CsvSchemaBuilder()
            .Column("Id").AsInt32()
            .Column("Amount").AsType(typeof(decimal))
            .Column("Active").AsBoolean()
            .Done()
            .Build();

        try
        {
            File.WriteAllText(path, "Id,Amount,Active\n-2147483648,-12.50,1\n");
            using var reader = CsvDocument.CreateDataReader(
                path,
                new CsvLoadOptions { Mode = CsvLoadMode.Stream },
                new CsvDataReaderOptions { Schema = schema });

            Assert.True(reader.Read());
            Assert.Equal(int.MinValue, reader.GetInt32(0));
            Assert.Equal(-12.50m, reader.GetDecimal(1));
            Assert.True(reader.GetBoolean(2));
            Assert.False(reader.Read());
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void CreateDataReader_FromCompressedFileWithExplicitSchema_ExposesTypedRows()
    {
        var path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.Tests." + Guid.NewGuid().ToString("N") + ".csv.gz");
        var schema = new CsvSchemaBuilder()
            .Column("Id").AsInt32()
            .Column("DisplayName").AsString()
            .Column("Score").AsType(typeof(decimal))
            .Column("CreatedUtc").AsDateTime()
            .Done()
            .Build();

        try
        {
            new CsvDocument()
                .WithHeader("Id", "DisplayName", "Score", "CreatedUtc")
                .AddRow(1, "Alice", 12.5m, new DateTime(2026, 1, 2, 3, 4, 5))
                .AddRow(2, "Bob", 13.75m, new DateTime(2026, 1, 3, 4, 5, 6))
                .Save(path, new CsvSaveOptions { CompressionType = CsvCompressionType.GZip });

            using var reader = CsvDocument.CreateDataReader(
                path,
                new CsvLoadOptions
                {
                    Mode = CsvLoadMode.Stream,
                    CompressionType = CsvCompressionType.GZip
                },
                new CsvDataReaderOptions { Schema = schema });

            Assert.Equal(typeof(int), reader.GetFieldType(reader.GetOrdinal("Id")));
            Assert.Equal(typeof(decimal), reader.GetFieldType(reader.GetOrdinal("Score")));

            Assert.True(reader.Read());
            Assert.Equal(1, reader.GetInt32(reader.GetOrdinal("Id")));
            Assert.Equal("Alice", reader.GetString(reader.GetOrdinal("DisplayName")));
            Assert.Equal(12.5m, reader.GetDecimal(reader.GetOrdinal("Score")));
            Assert.Equal(new DateTime(2026, 1, 2, 3, 4, 5), reader.GetDateTime(reader.GetOrdinal("CreatedUtc")));

            Assert.True(reader.Read());
            Assert.Equal(2, reader.GetInt32(reader.GetOrdinal("Id")));
            Assert.Equal("Bob", reader.GetString(reader.GetOrdinal("DisplayName")));
            Assert.Equal(13.75m, reader.GetDecimal(reader.GetOrdinal("Score")));
            Assert.Equal(new DateTime(2026, 1, 3, 4, 5, 6), reader.GetDateTime(reader.GetOrdinal("CreatedUtc")));
            Assert.False(reader.Read());
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void CreateDataReader_WithoutSchema_ExposesParsedValuesAsStrings()
    {
        var doc = CsvDocument.Parse("Id,Amount,Note\n1,12.5,\n");

        using var reader = doc.CreateDataReader();

        Assert.Equal(typeof(string), reader.GetFieldType(0));
        Assert.Equal(typeof(string), reader.GetFieldType(1));
        Assert.True(reader.Read());
        Assert.Equal("1", reader.GetString(0));
        Assert.Equal("12.5", reader.GetString(1));
        Assert.Equal(string.Empty, reader.GetString(2));
    }

    [Fact]
    public void GetChars_ReturnsZero_WhenOffsetIsPastField()
    {
        var doc = CsvDocument.Parse("Name\nAlpha\n");

        using var reader = doc.CreateDataReader();

        Assert.True(reader.Read());
        var buffer = new char[4];
        Assert.Equal(0, reader.GetChars(0, 99, buffer, 0, buffer.Length));
    }

    [Fact]
    public void CreateDataReader_WithoutSchema_ConvertsObjectValuesToStringsAndDbNull()
    {
        var doc = new CsvDocument()
            .WithHeader("Id", "Missing")
            .AddRow(42, null);

        using var reader = doc.CreateDataReader();

        Assert.True(reader.Read());
        Assert.Equal("42", reader.GetString(0));
        Assert.True(reader.IsDBNull(1));

        var values = new object[2];
        Assert.Equal(2, reader.GetValues(values));
        Assert.Equal("42", values[0]);
        Assert.Equal(DBNull.Value, values[1]);
    }

    [Fact]
    public void CreateDataReader_StreamingWithoutSchema_PreservesNullAndStaticColumns()
    {
        var doc = CsvDocument.Parse(
            "Id,Note\n1,<null>\n",
            new CsvLoadOptions
            {
                Mode = CsvLoadMode.Stream,
                NullValue = "<null>",
                StaticColumns = new Dictionary<string, object?> { ["Batch"] = 7 }
            });

        using var reader = doc.CreateDataReader();

        Assert.Equal(3, reader.FieldCount);
        Assert.True(reader.Read());
        Assert.Equal("1", reader.GetString(0));
        Assert.True(reader.IsDBNull(1));
        Assert.Equal("7", reader.GetString(2));

        var values = new object[reader.FieldCount];
        Assert.Equal(3, reader.GetValues(values));
        Assert.Equal("1", values[0]);
        Assert.Equal(DBNull.Value, values[1]);
        Assert.Equal("7", values[2]);
    }

    [Fact]
    public void CreateDataReader_InMemoryWithoutSchema_PreservesConfiguredNullValue()
    {
        var doc = CsvDocument.Parse(
            "Id,Note\n1,<null>\n",
            new CsvLoadOptions { NullValue = "<null>" });

        using var reader = doc.CreateDataReader();

        Assert.True(reader.Read());
        Assert.Equal("1", reader.GetString(0));
        Assert.True(reader.IsDBNull(1));

        var values = new object[reader.FieldCount];
        Assert.Equal(2, reader.GetValues(values));
        Assert.Equal("1", values[0]);
        Assert.Equal(DBNull.Value, values[1]);
    }

    [Fact]
    public void CreateDataReader_StreamingWithoutSchema_HonorsStrictColumnCounts()
    {
        var doc = CsvDocument.Parse(
            "First,Second\n1\n",
            new CsvLoadOptions
            {
                Mode = CsvLoadMode.Stream,
                ColumnCountMismatchPolicy = CsvColumnCountMismatchPolicy.Strict
            });

        using var reader = doc.CreateDataReader();

        var ex = Assert.Throws<CsvException>(() => reader.Read());
        Assert.Contains("Row contains 1 values but header defines 2 columns", ex.Message);
    }

    [Fact]
    public void CreateDataReader_WithRequiredSchema_RejectsMissingValues()
    {
        var doc = new CsvDocument()
            .WithHeader("Id")
            .AddRow(new object?[] { null });
        doc.EnsureSchema(schema => schema.Column("Id").AsInt32().Required());

        using var reader = doc.CreateDataReader();

        Assert.True(reader.Read());
        var ex = Assert.Throws<CsvException>(() => reader.GetValue(0));
        Assert.Contains("Column 'Id' is required", ex.Message);
    }

    [Fact]
    public void CreateDataReader_WithRequiredStringSchema_RejectsEmptyValues()
    {
        var doc = new CsvDocument()
            .WithHeader("Name")
            .AddRow(string.Empty);
        doc.EnsureSchema(schema => schema.Column("Name").AsString().Required());

        using var reader = doc.CreateDataReader();

        Assert.True(reader.Read());
        var ex = Assert.Throws<CsvException>(() => reader.GetValue(0));
        Assert.Contains("Column 'Name' is required", ex.Message);
    }
}
