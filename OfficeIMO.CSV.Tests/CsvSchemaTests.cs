using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvSchemaTests
{
    [Fact]
    public void MissingRequiredColumn_FailsValidation()
    {
        var doc = new CsvDocument()
            .WithHeader("Id", "Name")
            .AddRow(1, "Alice");

        doc.EnsureSchema(schema => schema
            .Column("Id").AsInt32().Required()
            .Column("Name").AsString().Required()
            .Column("Age").AsInt32().Required());

        doc.Validate(out var errors);
        Assert.NotEmpty(errors);
        Assert.Contains(errors, e => e.ColumnName == "Age");
    }

    [Fact]
    public void InvalidType_FailsValidation()
    {
        var doc = new CsvDocument()
            .WithHeader("Id", "Age")
            .AddRow(1, "abc");

        doc.EnsureSchema(schema => schema
            .Column("Id").AsInt32().Required()
            .Column("Age").AsInt32().Required());

        var ex = Assert.Throws<CsvValidationException>(() => doc.ValidateOrThrow());
        Assert.Contains(ex.Errors, e => e.ColumnName == "Age");
    }

    [Fact]
    public void InferSchema_Detects_Types_And_Required_Columns()
    {
        var doc = CsvDocument.Parse("Id,Amount,Active,Created,Note\n1,12.5,true,2026-07-07,Alpha\n2,13.7,false,2026-07-08,\n");

        var schema = doc.InferSchema();

        Assert.Equal(typeof(int), schema.Columns[0].DataType);
        Assert.Equal(typeof(decimal), schema.Columns[1].DataType);
        Assert.Equal(typeof(bool), schema.Columns[2].DataType);
        Assert.Equal(typeof(DateTime), schema.Columns[3].DataType);
        Assert.Equal(typeof(string), schema.Columns[4].DataType);
        Assert.True(schema.Columns[0].IsRequired);
        Assert.False(schema.Columns[4].IsRequired);
    }

    [Fact]
    public void InferSchema_Uses_Configured_DateTime_Formats()
    {
        var doc = CsvDocument.Parse(
            "Created\n07-Jul-2026\n",
            new CsvLoadOptions { DateTimeFormats = new[] { "dd-MMM-yyyy" } });

        var schema = doc.InferSchema();

        Assert.Equal(typeof(DateTime), Assert.Single(schema.Columns).DataType);
    }

    [Fact]
    public void InferSchema_ForTypedRows_IsNotOrderDependent()
    {
        var doc = new CsvDocument()
            .WithHeader("Value")
            .AddRow(1234567890123L)
            .AddRow(1);

        var schema = doc.InferSchema();

        Assert.Equal(typeof(long), Assert.Single(schema.Columns).DataType);
    }

    [Fact]
    public void EnsureInferredSchema_Attaches_Schema_For_Validation()
    {
        var doc = CsvDocument.Parse("Id,Name\n1,Alice\n2,Bob\n")
            .EnsureInferredSchema();

        doc.Validate(out var errors);

        Assert.Empty(errors);
    }

    [Fact]
    public void ToDataTable_WithInferredSchema_UsesTypedColumns()
    {
        var doc = CsvDocument.Parse("Id,Amount,Active,Created,Note\n1,12.5,true,2026-07-07,Alpha\n2,13.7,false,2026-07-08,\n");

        var table = doc.ToDataTable(new CsvDataTableOptions
        {
            TableName = "Rows",
            InferSchema = true
        });

        Assert.Equal("Rows", table.TableName);
        Assert.Equal(typeof(int), table.Columns["Id"]!.DataType);
        Assert.Equal(typeof(decimal), table.Columns["Amount"]!.DataType);
        Assert.Equal(typeof(bool), table.Columns["Active"]!.DataType);
        Assert.Equal(typeof(DateTime), table.Columns["Created"]!.DataType);
        Assert.Equal(typeof(string), table.Columns["Note"]!.DataType);
        Assert.Equal(1, table.Rows[0]["Id"]);
        Assert.Equal(12.5m, table.Rows[0]["Amount"]);
        Assert.Equal(false, table.Rows[1]["Active"]);
        Assert.Equal(string.Empty, table.Rows[1]["Note"]);
    }

    [Fact]
    public void ToDataTable_InStreamingMode_UsesTypedColumns()
    {
        var doc = CsvDocument.Parse(
            "Id,Amount,Active,Created,Note\n1,12.5,true,2026-07-07,Alpha\n2,13.7,false,2026-07-08,\n",
            new CsvLoadOptions { Mode = CsvLoadMode.Stream });

        var table = doc.ToDataTable(new CsvDataTableOptions
        {
            TableName = "Rows",
            InferSchema = true
        });

        Assert.Equal("Rows", table.TableName);
        Assert.Equal(typeof(int), table.Columns["Id"]!.DataType);
        Assert.Equal(typeof(decimal), table.Columns["Amount"]!.DataType);
        Assert.Equal(typeof(bool), table.Columns["Active"]!.DataType);
        Assert.Equal(typeof(DateTime), table.Columns["Created"]!.DataType);
        Assert.Equal(typeof(string), table.Columns["Note"]!.DataType);
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(1, table.Rows[0]["Id"]);
        Assert.Equal(13.7m, table.Rows[1]["Amount"]);
        Assert.Equal(string.Empty, table.Rows[1]["Note"]);
    }

    [Fact]
    public void ToDataTable_WithInferredSchema_StoresMissingTypedValuesAsDbNull()
    {
        var doc = CsvDocument.Parse("Id,Amount\n1,12.5\n2,\n");

        var table = doc.ToDataTable(new CsvDataTableOptions { InferSchema = true });

        Assert.Equal(typeof(decimal), table.Columns["Amount"]!.DataType);
        Assert.Same(DBNull.Value, table.Rows[1]["Amount"]);
    }

    [Fact]
    public void ToDataTable_WithRequiredSchema_RejectsMissingValues()
    {
        var doc = new CsvDocument()
            .WithHeader("Id")
            .AddRow(new object?[] { null });
        doc.EnsureSchema(schema => schema.Column("Id").AsInt32().Required());

        var ex = Assert.Throws<CsvException>(() => doc.ToDataTable());

        Assert.Contains("Column 'Id' is required", ex.Message);
    }

    [Fact]
    public void ToDataTable_InStreamingMode_PreservesNullAndStaticColumns()
    {
        var doc = CsvDocument.Parse(
            "Id,Note\n1,<null>\n",
            new CsvLoadOptions
            {
                Mode = CsvLoadMode.Stream,
                NullValue = "<null>",
                StaticColumns = new Dictionary<string, object?> { ["Batch"] = 7 }
            });

        var table = doc.ToDataTable(new CsvDataTableOptions { TableName = "Import" });

        Assert.Equal("Import", table.TableName);
        Assert.Equal(new[] { "Id", "Note", "Batch" }, table.Columns.Cast<DataColumn>().Select(column => column.ColumnName));
        Assert.Equal("1", table.Rows[0]["Id"]);
        Assert.Same(DBNull.Value, table.Rows[0]["Note"]);
        Assert.Equal("7", table.Rows[0]["Batch"]);
    }

    [Fact]
    public void ToDataTable_InStreamingMode_HonorsStrictColumnCounts()
    {
        var doc = CsvDocument.Parse(
            "First,Second\n1\n",
            new CsvLoadOptions
            {
                Mode = CsvLoadMode.Stream,
                ColumnCountMismatchPolicy = CsvColumnCountMismatchPolicy.Strict
            });

        var ex = Assert.Throws<CsvException>(() => doc.ToDataTable());

        Assert.Contains("Row contains 1 values but header defines 2 columns", ex.Message);
    }

    [Fact]
    public void Malformed_Unterminated_Quote_Throws()
    {
        var csv = "Id,Name\n1,\"Unclosed";
        Assert.Throws<CsvParseException>(() => CsvDocument.Parse(csv));
    }
}
