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
    public void Malformed_Unterminated_Quote_Throws()
    {
        var csv = "Id,Name\n1,\"Unclosed";
        Assert.Throws<CsvParseException>(() => CsvDocument.Parse(csv));
    }
}
