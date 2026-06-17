using System.Globalization;
using System.Linq;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvDocumentBasicsTests
{
    [Fact]
    public void RoundTrip_ToString_ParsesBack()
    {
        var doc = new CsvDocument()
            .WithHeader("Name", "Age", "City");

        doc.AddRow("Przemek", 36, "Mikołów")
           .AddRow("Dominika", 30, "Mikołów");

        var text = doc.ToString();
        var parsed = CsvDocument.Parse(text);

        Assert.Equal(doc.Header, parsed.Header);
        Assert.Equal(2, parsed.AsEnumerable().Count());
        Assert.Equal("Przemek", parsed.AsEnumerable().ElementAt(0).AsString("Name"));
        Assert.Equal(36, parsed.AsEnumerable().ElementAt(0).AsInt32("Age"));
    }

    [Fact]
    public void Supports_Custom_Delimiters()
    {
        var doc = new CsvDocument()
            .WithDelimiter(';')
            .WithHeader("Name", "Age");

        doc.AddRow("Ala", 10)
           .AddRow("Ola", 12);

        var text = doc.ToString();
        var options = new CsvLoadOptions { Delimiter = ';' };
        var parsed = CsvDocument.Parse(text, options);

        Assert.Equal(';', parsed.Delimiter);
        Assert.Equal(2, parsed.AsEnumerable().Count());
        Assert.Equal(12, parsed.AsEnumerable().ElementAt(1).AsInt32("Age"));
    }

    [Theory]
    [InlineData(',')]
    [InlineData(';')]
    [InlineData('\t')]
    public void Handles_Quoted_Fields_With_Delimiters(char delimiter)
    {
        var doc = new CsvDocument()
            .WithDelimiter(delimiter)
            .WithHeader("Id", "Value");

        doc.AddRow(1, $"hello{delimiter}world");
        var text = doc.ToString();

        var parsed = CsvDocument.Parse(text, new CsvLoadOptions { Delimiter = delimiter });
        var value = parsed.AsEnumerable().Single().AsString("Value");
        Assert.Equal($"hello{delimiter}world", value);
    }

    [Fact]
    public void Empty_File_Produces_Empty_Document()
    {
        var parsed = CsvDocument.Parse(string.Empty);
        Assert.Empty(parsed.Header);
        Assert.Empty(parsed.AsEnumerable());
    }

    [Fact]
    public void Header_Only_Has_No_Rows()
    {
        var csv = "Name,Age\n";
        var parsed = CsvDocument.Parse(csv);
        Assert.Equal(new[] { "Name", "Age" }, parsed.Header);
        Assert.Empty(parsed.AsEnumerable());
    }

    [Fact]
    public void Formula_Injection_Policy_Preserves_Values_By_Default()
    {
        var doc = new CsvDocument()
            .WithHeader("=Header");

        doc.AddRow("=cmd");

        var text = doc.ToString(new CsvSaveOptions { NewLine = "\n" });

        Assert.Equal("=Header\n=cmd\n", text);
    }

    [Theory]
    [InlineData("=cmd", "'=cmd")]
    [InlineData("+cmd", "'+cmd")]
    [InlineData("-cmd", "'-cmd")]
    [InlineData("@cmd", "'@cmd")]
    [InlineData("\tcmd", "'\tcmd")]
    [InlineData("  =cmd", "'  =cmd")]
    public void Formula_Injection_Policy_Escapes_Dangerous_Row_Values(string value, string expected)
    {
        var doc = new CsvDocument()
            .WithHeader("Value");

        doc.AddRow(value);

        var text = doc.ToString(new CsvSaveOptions {
            NewLine = "\n",
            FormulaInjectionPolicy = CsvFormulaInjectionPolicy.Escape
        });

        Assert.Equal($"Value\n{expected}\n", text);
    }

    [Fact]
    public void Formula_Injection_Policy_Escapes_Dangerous_Headers()
    {
        var doc = new CsvDocument()
            .WithHeader("=Header");

        doc.AddRow("safe");

        var text = doc.ToString(new CsvSaveOptions {
            NewLine = "\n",
            FormulaInjectionPolicy = CsvFormulaInjectionPolicy.Escape
        });

        Assert.Equal("'=Header\nsafe\n", text);
    }
}
