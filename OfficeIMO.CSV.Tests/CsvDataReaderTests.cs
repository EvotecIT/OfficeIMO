using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
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
}
