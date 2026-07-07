using System;
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
