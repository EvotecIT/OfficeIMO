using System;
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
    public void Malformed_Unterminated_Quote_Throws()
    {
        var csv = "Id,Name\n1,\"Unclosed";
        Assert.Throws<CsvParseException>(() => CsvDocument.Parse(csv));
    }
}
