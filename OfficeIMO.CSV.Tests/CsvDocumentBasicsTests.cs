using System.Globalization;
using System.IO;
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

    [Fact]
    public void Detects_Delimiter_From_Header_Even_When_Data_Contains_More_Commas()
    {
        var parsed = CsvDocument.Parse(
            "Field1;Field2;Field3\n1,2,3,4;5,6,7,8;9,10,11,12\n",
            new CsvLoadOptions { DetectDelimiter = true });

        Assert.Equal(';', parsed.Delimiter);
        Assert.Equal(new[] { "Field1", "Field2", "Field3" }, parsed.Header);
        Assert.Equal("5,6,7,8", parsed.AsEnumerable().Single().AsString("Field2"));
    }

    [Fact]
    public void Detects_Delimiter_With_Custom_Candidates_In_Streaming_Mode()
    {
        var parsed = CsvDocument.Parse(
            "Name|Value\nAlpha|1\n",
            new CsvLoadOptions {
                DetectDelimiter = true,
                DelimiterCandidates = new[] { ';', '|' },
                Mode = CsvLoadMode.Stream
            });

        Assert.Equal('|', parsed.Delimiter);
        Assert.Equal("1", parsed.AsEnumerable().Single().AsString("Value"));
    }

    [Fact]
    public void Preserves_Unquoted_Whitespace_By_Default()
    {
        var parsed = CsvDocument.Parse("Name,Value\nAlpha,  spaced  \n");

        Assert.Equal("  spaced  ", parsed.AsEnumerable().Single().AsString("Value"));
    }

    [Fact]
    public void Can_Trim_Unquoted_Whitespace_When_Requested()
    {
        var parsed = CsvDocument.Parse(
            "Name,Value\nAlpha,  spaced  \n",
            new CsvLoadOptions { TrimWhitespace = true });

        Assert.Equal("spaced", parsed.AsEnumerable().Single().AsString("Value"));
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
    public void Explicit_Header_Treats_First_Record_As_Data()
    {
        var parsed = CsvDocument.Parse(
            "Alpha,1\nBeta,2\n",
            new CsvLoadOptions { Header = new[] { "Name", "Value" } });

        Assert.Equal(new[] { "Name", "Value" }, parsed.Header);
        Assert.Equal(2, parsed.AsEnumerable().Count());
        Assert.Equal("Alpha", parsed.AsEnumerable().ElementAt(0).AsString("Name"));
    }

    [Fact]
    public void Missing_Header_Names_Are_Generated_By_Default()
    {
        var parsed = CsvDocument.Parse("Name,,Value\nAlpha,Ignored,1\n");

        Assert.Equal(new[] { "Name", "H1", "Value" }, parsed.Header);
        Assert.Equal("Ignored", parsed.AsEnumerable().Single().AsString("H1"));
    }

    [Fact]
    public void Parsed_Rows_Pad_Missing_Fields_And_Ignore_Extras_By_Default()
    {
        var parsed = CsvDocument.Parse("Name,Value\nAlpha\nBeta,2,Extra\n");
        var rows = parsed.AsEnumerable().ToArray();

        Assert.Equal(string.Empty, rows[0].AsString("Value"));
        Assert.Equal("2", rows[1].AsString("Value"));
        Assert.Equal(2, rows[1].FieldCount);
    }

    [Fact]
    public void Strict_Column_Count_Mismatch_Policy_Throws()
    {
        var ex = Assert.Throws<CsvException>(() => CsvDocument.Parse(
            "Name,Value\nAlpha\n",
            new CsvLoadOptions { ColumnCountMismatchPolicy = CsvColumnCountMismatchPolicy.Strict }));

        Assert.Contains("header defines 2 columns", ex.Message);
    }

    [Fact]
    public void Streaming_Mode_Uses_Column_Count_Mismatch_Policy()
    {
        var parsed = CsvDocument.Parse(
            "Name,Value\nAlpha\nBeta,2,Extra\n",
            new CsvLoadOptions { Mode = CsvLoadMode.Stream });
        var rows = parsed.AsEnumerable().ToArray();

        Assert.Equal(string.Empty, rows[0].AsString("Value"));
        Assert.Equal("2", rows[1].AsString("Value"));
        Assert.Equal(2, rows[1].FieldCount);
    }

    [Fact]
    public void Skips_Comments_Before_Header_By_Default()
    {
        var parsed = CsvDocument.Parse("#Version: 1.0\nName,Value\nAlpha,1\n");

        Assert.Equal(new[] { "Name", "Value" }, parsed.Header);
        Assert.Single(parsed.AsEnumerable());
        Assert.Equal("Alpha", parsed.AsEnumerable().Single().AsString("Name"));
    }

    [Fact]
    public void Can_Treat_Leading_Comment_As_Header_When_Requested()
    {
        var parsed = CsvDocument.Parse(
            "#Name,Value\nAlpha,1\n",
            new CsvLoadOptions { SkipCommentRowsBeforeHeader = false });

        Assert.Equal(new[] { "#Name", "Value" }, parsed.Header);
        Assert.Single(parsed.AsEnumerable());
        Assert.Equal("Alpha", parsed.AsEnumerable().Single()["#Name"]);
    }

    [Fact]
    public void Can_Skip_Comment_Rows_Throughout_File()
    {
        var parsed = CsvDocument.Parse(
            "Name,Value\nAlpha,1\n# ignored\nBeta,2\n",
            new CsvLoadOptions { SkipCommentRows = true });

        var rows = parsed.AsEnumerable().ToArray();

        Assert.Equal(2, rows.Length);
        Assert.Equal("Alpha", rows[0].AsString("Name"));
        Assert.Equal("Beta", rows[1].AsString("Name"));
    }

    [Fact]
    public void Can_Skip_Custom_Comment_Rows_Throughout_File()
    {
        var parsed = CsvDocument.Parse(
            "Name,Value\nAlpha,1\n; ignored\nBeta,2\n",
            new CsvLoadOptions {
                SkipCommentRows = true,
                CommentCharacter = ';'
            });

        var rows = parsed.AsEnumerable().ToArray();

        Assert.Equal(2, rows.Length);
        Assert.Equal("Beta", rows[1].AsString("Name"));
    }

    [Fact]
    public void Recognizes_W3C_Fields_Header()
    {
        var parsed = CsvDocument.Parse(
            "#Version: 1.0\n#Fields: date time cs-uri\n2026-06-24 12:00 /index\n",
            new CsvLoadOptions { Delimiter = ' ' });

        Assert.Equal(new[] { "date", "time", "cs-uri" }, parsed.Header);
        Assert.Single(parsed.AsEnumerable());
        Assert.Equal("/index", parsed.AsEnumerable().Single().AsString("cs-uri"));
    }

    [Fact]
    public void Skip_Comment_Rows_Does_Not_Skip_W3C_Fields_Header()
    {
        var parsed = CsvDocument.Parse(
            "#Version: 1.0\n#Fields: date time cs-uri\n#Software: test\n2026-06-24 12:00 /index\n",
            new CsvLoadOptions {
                Delimiter = ' ',
                SkipCommentRows = true
            });

        Assert.Equal(new[] { "date", "time", "cs-uri" }, parsed.Header);
        Assert.Single(parsed.AsEnumerable());
        Assert.Equal("/index", parsed.AsEnumerable().Single().AsString("cs-uri"));
    }

    [Fact]
    public void Streaming_Mode_Uses_Same_Header_Discovery()
    {
        var parsed = CsvDocument.Parse(
            "#Version: 1.0\n#Fields: date time cs-uri\n2026-06-24 12:00 /index\n",
            new CsvLoadOptions { Delimiter = ' ', Mode = CsvLoadMode.Stream });

        Assert.Equal(new[] { "date", "time", "cs-uri" }, parsed.Header);
        Assert.Single(parsed.AsEnumerable());
        Assert.Equal("2026-06-24", parsed.AsEnumerable().Single().AsString("date"));
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

    [Fact]
    public void Formula_Injection_Policy_Escapes_Span_Formatted_Row_Values()
    {
        using var writer = new StringWriter();
        using (var csv = new CsvObjectWriter(writer, new CsvSaveOptions {
            NewLine = "\n",
            FormulaInjectionPolicy = CsvFormulaInjectionPolicy.Escape
        }))
        {
            csv.WriteRow(new[] { "Value" }, new object?[] { -1 });
        }

        Assert.Equal("Value\n'-1\n", writer.ToString());
    }
}
