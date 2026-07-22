using System;
using System.Globalization;
using System.IO;
using System.Linq;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvDocumentBasicsTests
{
    [Fact]
    public void Quoted_Field_Delimiters_Do_Not_Inflate_Field_Allocation()
    {
        string value = new string(',', 100_000);

        CsvDocument parsed = CsvDocument.Parse(
            "\"" + value + "\"\n",
            new CsvLoadOptions { HasHeaderRow = false });

        Assert.Equal(value, parsed.AsEnumerable().Single().AsString("Column1"));
    }

    [Fact]
    public void Multiline_Quoted_Record_Is_Parsed_Without_Reparsing_Each_Prefix()
    {
        string value = string.Join("\n", Enumerable.Repeat("payload", 5_000));

        CsvDocument parsed = CsvDocument.Parse(
            "\"" + value + "\",done\n",
            new CsvLoadOptions { HasHeaderRow = false });

        CsvRow row = parsed.AsEnumerable().Single();
        Assert.Equal(value, row.AsString("Column1"));
        Assert.Equal("done", row.AsString("Column2"));
    }

    [Fact]
    public void Unterminated_Multiline_Quote_Fails_After_A_Single_Forward_Scan()
    {
        string value = "\"" + string.Join("\n", Enumerable.Repeat("payload", 5_000));

        Assert.Throws<CsvParseException>(() => CsvDocument.Parse(
            value,
            new CsvLoadOptions { HasHeaderRow = false }));
    }

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

    [Fact]
    public void TrimWhitespace_Preserves_Quoted_Whitespace()
    {
        var parsed = CsvDocument.Parse(
            "\"Name\",\"Value\"\n\"Alpha\",\"  spaced  \"\n",
            new CsvLoadOptions { TrimWhitespace = true });

        Assert.Equal("  spaced  ", parsed.AsEnumerable().Single().AsString("Value"));
    }

    [Fact]
    public void Preserves_Surrounding_Whitespace_Around_Quoted_Fields_By_Default()
    {
        var parsed = CsvDocument.Parse("Name\n  \"Alpha\"  \n");

        Assert.Equal("  Alpha  ", parsed.AsEnumerable().Single().AsString("Name"));
    }

    [Fact]
    public void Delimiter_Detection_Skips_Trimmed_Blank_Records()
    {
        var parsed = CsvDocument.Parse(
            "   \nName;Value\nAlpha;1\n",
            new CsvLoadOptions
            {
                DetectDelimiter = true,
                TrimWhitespace = true
            });

        var row = Assert.Single(parsed.AsEnumerable());
        Assert.Equal(';', parsed.Delimiter);
        Assert.Equal("Alpha", row.AsString("Name"));
        Assert.Equal("1", row.AsString("Value"));
    }

    [Fact]
    public void Delimiter_Detection_Applies_Skip_Before_Preserved_Blank_Records()
    {
        var parsed = CsvDocument.Parse(
            "\nName;Value\nAlpha;1\n",
            new CsvLoadOptions
            {
                DetectDelimiter = true,
                AllowEmptyLines = true,
                SkipInitialRecords = 1
            });

        var row = Assert.Single(parsed.AsEnumerable());
        Assert.Equal(';', parsed.Delimiter);
        Assert.Equal(new[] { "Name", "Value" }, parsed.Header);
        Assert.Equal("Alpha", row.AsString("Name"));
        Assert.Equal("1", row.AsString("Value"));
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

    [Theory]
    [InlineData("Alpha, \"Beta\" ", " Beta ")]
    [InlineData("Alpha, \"Be,ta\" ", " Be,ta ")]
    [InlineData("Alpha,\"Be\"  ", "Be  ")]
    public void Preserves_Padding_Around_Quoted_Fields_By_Default(string row, string expected)
    {
        var parsed = CsvDocument.Parse($"Name,Value\n{row}\n");

        Assert.Equal(expected, parsed.AsEnumerable().Single().AsString("Value"));
    }

    [Theory]
    [InlineData("Alpha, \"Beta\" ", "Beta")]
    [InlineData("Alpha, \"Be,ta\" ", "Be,ta")]
    [InlineData("Alpha,\"Be\"  ", "Be")]
    public void TrimWhitespace_Removes_Padding_Around_Quoted_Fields(string row, string expected)
    {
        var parsed = CsvDocument.Parse(
            $"Name,Value\n{row}\n",
            new CsvLoadOptions { TrimWhitespace = true });

        Assert.Equal(expected, parsed.AsEnumerable().Single().AsString("Value"));
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
    public void Blank_Lines_Are_Skipped_By_Default()
    {
        var parsed = CsvDocument.Parse("Name,Value\n\nAlpha,1\n");

        var row = Assert.Single(parsed.AsEnumerable());
        Assert.Equal("Alpha", row.AsString("Name"));
    }

    [Fact]
    public void Delimiter_Only_Rows_Are_Data_Rows()
    {
        var parsed = CsvDocument.Parse("Name,Value\n,\nAlpha,1\n");
        var rows = parsed.AsEnumerable().ToArray();

        Assert.Equal(2, rows.Length);
        Assert.Equal(string.Empty, rows[0].AsString("Name"));
        Assert.Equal(string.Empty, rows[0].AsString("Value"));
        Assert.Equal("Alpha", rows[1].AsString("Name"));
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
    public void Can_Skip_Initial_Records_Before_Header_Discovery()
    {
        var parsed = CsvDocument.Parse(
            "generated by vendor\nexported 2026-06-25\nName,Value\nAlpha,1\n",
            new CsvLoadOptions { SkipInitialRecords = 2 });

        var row = Assert.Single(parsed.AsEnumerable());
        Assert.Equal(new[] { "Name", "Value" }, parsed.Header);
        Assert.Equal("Alpha", row.AsString("Name"));
    }

    [Fact]
    public void Delimiter_Detection_Skips_Initial_Records()
    {
        var parsed = CsvDocument.Parse(
            "generated,by,vendor,with,commas\nName;Value\nAlpha;1\n",
            new CsvLoadOptions {
                DetectDelimiter = true,
                SkipInitialRecords = 1
            });

        var row = Assert.Single(parsed.AsEnumerable());
        Assert.Equal(';', parsed.Delimiter);
        Assert.Equal(new[] { "Name", "Value" }, parsed.Header);
        Assert.Equal("1", row.AsString("Value"));
    }

    [Fact]
    public void Delimiter_Detection_Skips_Initial_Records_After_Leading_Comments()
    {
        var parsed = CsvDocument.Parse(
            "#note\nmetadata,with,commas\nName;Value\nAlpha;1\n",
            new CsvLoadOptions {
                DetectDelimiter = true,
                SkipInitialRecords = 1
            });

        var row = Assert.Single(parsed.AsEnumerable());
        Assert.Equal(';', parsed.Delimiter);
        Assert.Equal(new[] { "Name", "Value" }, parsed.Header);
        Assert.Equal("Alpha", row.AsString("Name"));
        Assert.Equal("1", row.AsString("Value"));
    }

    [Fact]
    public void Skip_Initial_Records_Applies_Before_W3C_Header_Recognition()
    {
        var parsed = CsvDocument.Parse(
            "#Fields: Old Value\nName,Value\nAlpha,1\n",
            new CsvLoadOptions { SkipInitialRecords = 1 });

        var row = Assert.Single(parsed.AsEnumerable());
        Assert.Equal(new[] { "Name", "Value" }, parsed.Header);
        Assert.Equal("Alpha", row.AsString("Name"));
        Assert.Equal("1", row.AsString("Value"));
    }

    [Fact]
    public void Skip_Initial_Records_Keeps_W3C_Header_Recognition_When_Comments_Are_Skipped()
    {
        var parsed = CsvDocument.Parse(
            "metadata\n#Fields: date time\n2026-06-25 12:00\n",
            new CsvLoadOptions
            {
                Delimiter = ' ',
                SkipInitialRecords = 1,
                SkipCommentRows = true
            });

        var row = Assert.Single(parsed.AsEnumerable());

        Assert.Equal(new[] { "date", "time" }, parsed.Header);
        Assert.Equal("2026-06-25", row.AsString("date"));
        Assert.Equal("12:00", row.AsString("time"));
    }

    [Fact]
    public void Skip_Comment_Rows_Before_Header_Skips_Unmatched_Quote_Comments_Before_Parsing()
    {
        var parsed = CsvDocument.Parse(
            "# generated \"by tool\nName,Value\nA,1\n",
            new CsvLoadOptions());

        var row = Assert.Single(parsed.AsEnumerable());
        Assert.Equal(new[] { "Name", "Value" }, parsed.Header);
        Assert.Equal("A", row.AsString("Name"));
        Assert.Equal("1", row.AsString("Value"));
    }

    [Fact]
    public void Skip_Comment_Rows_Before_Header_Keeps_Delimiterless_Header_After_Unmatched_Quote_Comment()
    {
        var parsed = CsvDocument.Parse(
            "# generated \"by tool\nName\nAlpha\n",
            new CsvLoadOptions());

        var row = Assert.Single(parsed.AsEnumerable());
        Assert.Equal(new[] { "Name" }, parsed.Header);
        Assert.Equal("Alpha", row.AsString("Name"));
    }

    [Fact]
    public void Delimiter_Detection_Skips_Unmatched_Quote_Comments_Before_Sampling()
    {
        var parsed = CsvDocument.Parse(
            "# generated \"by tool\nName;Value\nA;1\n",
            new CsvLoadOptions { DetectDelimiter = true });

        var row = Assert.Single(parsed.AsEnumerable());
        Assert.Equal(';', parsed.Delimiter);
        Assert.Equal(new[] { "Name", "Value" }, parsed.Header);
        Assert.Equal("A", row.AsString("Name"));
        Assert.Equal("1", row.AsString("Value"));
    }

    [Fact]
    public void Skip_Comment_Rows_Before_Header_Skips_Delimiterless_Multiline_Comment_Record()
    {
        var parsed = CsvDocument.Parse(
            "#note \"ignored\nstill ignored\"\nName,Value\nA,1\n",
            new CsvLoadOptions());

        var row = Assert.Single(parsed.AsEnumerable());
        Assert.Equal(new[] { "Name", "Value" }, parsed.Header);
        Assert.Equal("A", row.AsString("Name"));
        Assert.Equal("1", row.AsString("Value"));
    }

    [Fact]
    public void Skip_Comment_Rows_Before_Header_Skips_Delimiterless_Comment_Closing_After_Multiple_Continuation_Lines()
    {
        var parsed = CsvDocument.Parse(
            "#note \"one\ntwo\nthree\"\nName,Value\nA,1\n",
            new CsvLoadOptions());

        var row = Assert.Single(parsed.AsEnumerable());
        Assert.Equal(new[] { "Name", "Value" }, parsed.Header);
        Assert.Equal("A", row.AsString("Name"));
        Assert.Equal("1", row.AsString("Value"));
    }

    [Fact]
    public void Skip_Comment_Rows_Before_Header_Skips_Complete_Multiline_Quoted_Comment_Record()
    {
        var parsed = CsvDocument.Parse(
            "#note,\"one\ntwo\nthree\"\nName,Value\nA,1\n",
            new CsvLoadOptions());

        var row = Assert.Single(parsed.AsEnumerable());
        Assert.Equal(new[] { "Name", "Value" }, parsed.Header);
        Assert.Equal("A", row.AsString("Name"));
        Assert.Equal("1", row.AsString("Value"));
    }

    [Fact]
    public void Delimiter_Detection_Skips_Complete_Multiline_Quoted_Comment_Record()
    {
        var parsed = CsvDocument.Parse(
            "#note,\"a\nstill,has,commas\nend\"\nName;Value\nA;1\n",
            new CsvLoadOptions { DetectDelimiter = true });

        var row = Assert.Single(parsed.AsEnumerable());
        Assert.Equal(';', parsed.Delimiter);
        Assert.Equal(new[] { "Name", "Value" }, parsed.Header);
        Assert.Equal("A", row.AsString("Name"));
        Assert.Equal("1", row.AsString("Value"));
    }

    [Fact]
    public void Delimiter_Detection_Skips_Delimiterless_Multiline_Comment_Record_With_Candidate_Delimiters()
    {
        var parsed = CsvDocument.Parse(
            "#note \"ignored\nstill,has,commas\"\nName;Value\nA;1\n",
            new CsvLoadOptions { DetectDelimiter = true });

        var row = Assert.Single(parsed.AsEnumerable());
        Assert.Equal(';', parsed.Delimiter);
        Assert.Equal(new[] { "Name", "Value" }, parsed.Header);
        Assert.Equal("A", row.AsString("Name"));
        Assert.Equal("1", row.AsString("Value"));
    }

    [Fact]
    public void Delimiter_Detection_Keeps_Comment_Looking_Data_When_No_Header_Is_Discovered()
    {
        var parsed = CsvDocument.Parse(
            "#a,b,c\n1;2;3\n4;5;6\n",
            new CsvLoadOptions {
                DetectDelimiter = true,
                HasHeaderRow = false
            });

        var rows = parsed.AsEnumerable().ToArray();
        Assert.Equal(',', parsed.Delimiter);
        Assert.Equal(new[] { "Column1", "Column2", "Column3" }, parsed.Header);
        Assert.Equal("#a", rows[0].AsString("Column1"));
        Assert.Equal("b", rows[0].AsString("Column2"));
        Assert.Equal("1;2;3", rows[1].AsString("Column1"));
    }

    [Fact]
    public void Delimiter_Detection_Keeps_Comment_Looking_Data_With_Explicit_Header()
    {
        var parsed = CsvDocument.Parse(
            "#Alpha,1\nBeta;2\n",
            new CsvLoadOptions {
                DetectDelimiter = true,
                Header = new[] { "Name", "Value" }
            });

        var rows = parsed.AsEnumerable().ToArray();
        Assert.Equal(',', parsed.Delimiter);
        Assert.Equal("#Alpha", rows[0].AsString("Name"));
        Assert.Equal("1", rows[0].AsString("Value"));
        Assert.Equal("Beta;2", rows[1].AsString("Name"));
    }

    [Fact]
    public void Delimiter_Detection_Keeps_Post_Header_Comment_Looking_Data()
    {
        var parsed = CsvDocument.Parse(
            "A,B;C\n#x,y,z\n1;2;3\n",
            new CsvLoadOptions { DetectDelimiter = true });

        var rows = parsed.AsEnumerable().ToArray();
        Assert.Equal(',', parsed.Delimiter);
        Assert.Equal(new[] { "A", "B;C" }, parsed.Header);
        Assert.Equal("#x", rows[0].AsString("A"));
        Assert.Equal("y", rows[0].AsString("B;C"));
        Assert.Equal("1;2;3", rows[1].AsString("A"));
    }

    [Fact]
    public void ReadRows_Preserves_Post_Header_Comment_Prefixed_Data()
    {
        var rows = new System.Collections.Generic.List<string>();
        using var reader = new StringReader("Name,Value\n#tag,1\nAlpha,2\n");

        CsvDocument.ReadRows(reader, (header, row) =>
        {
            Assert.Equal(new[] { "Name", "Value" }, header);
            rows.Add($"{row[0]}|{row[1]}");
        });

        Assert.Equal(new[] { "#tag|1", "Alpha|2" }, rows);
    }

    [Fact]
    public void ReadRowsReusable_Preserves_Post_Header_Comment_Prefixed_Data()
    {
        var rows = new System.Collections.Generic.List<string>();
        using var reader = new StringReader("Name,Value\n#tag,1\nAlpha,2\n");

        CsvDocument.ReadRowsReusable(reader, (header, row) =>
        {
            Assert.Equal(new[] { "Name", "Value" }, header);
            rows.Add($"{row[0]}|{row[1]}");
        });

        Assert.Equal(new[] { "#tag|1", "Alpha|2" }, rows);
    }

    [Fact]
    public void Can_Skip_Initial_Records_With_Explicit_Header()
    {
        var parsed = CsvDocument.Parse(
            "metadata\nAlpha,1\nBeta,2\n",
            new CsvLoadOptions {
                Header = new[] { "Name", "Value" },
                SkipInitialRecords = 1
            });

        var rows = parsed.AsEnumerable().ToArray();

        Assert.Equal(2, rows.Length);
        Assert.Equal("Alpha", rows[0].AsString("Name"));
        Assert.Equal("Beta", rows[1].AsString("Name"));
    }

    [Fact]
    public void Streaming_Mode_Can_Skip_Initial_Records()
    {
        var parsed = CsvDocument.Parse(
            "metadata\nName,Value\nAlpha,1\nBeta,2\n",
            new CsvLoadOptions {
                Mode = CsvLoadMode.Stream,
                SkipInitialRecords = 1
            });

        var rows = parsed.AsEnumerable().ToArray();

        Assert.Equal(new[] { "Name", "Value" }, parsed.Header);
        Assert.Equal(2, rows.Length);
        Assert.Equal("Beta", rows[1].AsString("Name"));
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
    public void Quoted_Comment_Character_Header_Is_Not_Treated_As_Comment()
    {
        var parsed = CsvDocument.Parse("\"#Tag\",Name\n10,Alpha\n");

        var row = Assert.Single(parsed.AsEnumerable());

        Assert.Equal(new[] { "#Tag", "Name" }, parsed.Header);
        Assert.Equal("10", row.AsString("#Tag"));
        Assert.Equal("Alpha", row.AsString("Name"));
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
    public void Skip_Comment_Rows_Skips_W3C_Markers_After_Normal_Header()
    {
        var parsed = CsvDocument.Parse(
            "Name,Value\n#Fields: old value\nAlpha,1\n",
            new CsvLoadOptions { SkipCommentRows = true });

        var row = Assert.Single(parsed.AsEnumerable());

        Assert.Equal(new[] { "Name", "Value" }, parsed.Header);
        Assert.Equal("Alpha", row.AsString("Name"));
        Assert.Equal("1", row.AsString("Value"));
    }

    [Fact]
    public void Delimiter_Detection_Skips_W3C_Markers_After_Normal_Header()
    {
        var parsed = CsvDocument.Parse(
            "A,B;C\n#Fields: old,value\n1;2;3\n",
            new CsvLoadOptions
            {
                DetectDelimiter = true,
                SkipCommentRows = true
            });

        Assert.Equal(';', parsed.Delimiter);
        Assert.Equal(new[] { "A,B", "C" }, parsed.Header);
    }

    [Fact]
    public void Delimiter_Detection_Skips_W3C_Markers_Only_When_Header_Can_Consume_Them()
    {
        var parsed = CsvDocument.Parse(
            "#Fields: old,value\n1;2\n3;4\n",
            new CsvLoadOptions
            {
                DetectDelimiter = true,
                HasHeaderRow = false,
                SkipCommentRows = true
            });

        Assert.Equal(';', parsed.Delimiter);
        Assert.Equal(2, parsed.AsEnumerable().Count());
    }

    [Fact]
    public void Skip_Comment_Rows_Skips_Complete_Multiline_Comment_Record()
    {
        var parsed = CsvDocument.Parse(
            "Name,Value\n#note,\"ignored\nstill ignored\"\nAlpha,1\n",
            new CsvLoadOptions { SkipCommentRows = true });

        var row = Assert.Single(parsed.AsEnumerable());

        Assert.Equal("Alpha", row.AsString("Name"));
        Assert.Equal("1", row.AsString("Value"));
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
    public void Recognizes_W3C_Fields_Header_With_Repeated_Whitespace()
    {
        var parsed = CsvDocument.Parse(
            "#Fields: date  time cs-uri\n2026-06-24 12:00 /index\n",
            new CsvLoadOptions { Delimiter = ' ' });

        var row = Assert.Single(parsed.AsEnumerable());
        Assert.Equal(new[] { "date", "time", "cs-uri" }, parsed.Header);
        Assert.Equal("2026-06-24", row.AsString("date"));
        Assert.Equal("12:00", row.AsString("time"));
        Assert.Equal("/index", row.AsString("cs-uri"));
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

    [Fact]
    public void Formula_Injection_Policy_Escapes_Buffered_Null_Tokens()
    {
        using var writer = new StringWriter();
        using (var csv = new CsvObjectWriter(writer, new CsvSaveOptions {
            NewLine = "\n",
            NullValue = "=NULL",
            FormulaInjectionPolicy = CsvFormulaInjectionPolicy.Escape
        }))
        {
            csv.WriteRow(new[] { "Value" }, new object?[] { null });
        }

        Assert.Equal("Value\n'=NULL\n", writer.ToString());
    }

    [Fact]
    public void Quoted_Multiline_Fields_Preserve_Physical_Newlines()
    {
        var parsed = CsvDocument.Parse("Name,Note\r\nAlpha,\"one\r\ntwo\"\r\n");

        Assert.Equal("one\r\ntwo", parsed.AsEnumerable().Single().AsString("Note"));
    }

    [Fact]
    public void Quoted_Multiline_Fields_Preserve_Lf_Newlines()
    {
        var parsed = CsvDocument.Parse("Name,Note\nAlpha,\"one\ntwo\"\n");
        var row = parsed.AsEnumerable().Single();

        Assert.Equal("one\ntwo", row[1]?.ToString());
        Assert.Equal("one\ntwo", row.AsString("Note"));
    }

    [Fact]
    public void Quoted_Multiline_Fields_Preserve_Lf_Newlines_At_End_Of_File()
    {
        var text = "Name,Note\nAlpha,\"one\ntwo\"";

        var parsed = CsvDocument.Parse(text);
        Assert.Equal("one\ntwo", parsed.AsEnumerable().Single().AsString("Note"));

        using var reader = new StringReader(text);
        string? streamedValue = null;
        CsvDocument.ReadRowsReusable(reader, (_, row) => streamedValue = row[1]);

        Assert.Equal("one\ntwo", streamedValue);
    }

#if NET6_0_OR_GREATER
    [Fact]
    public void Default_Writer_Preserves_Long_Culture_Formatted_Dates()
    {
        var culture = (CultureInfo)CultureInfo.InvariantCulture.Clone();
        culture.DateTimeFormat.ShortDatePattern = new string('y', 160);
        culture.DateTimeFormat.LongTimePattern = "HH:mm:ss";
        var value = new DateTime(2026, 7, 15, 12, 34, 56);
        using var writer = new StringWriter(culture);

        using (var csv = new CsvObjectWriter(writer, new CsvSaveOptions { Culture = culture, NewLine = "\n" }))
        {
            csv.WriteRow(new[] { "Value" }, new object?[] { value });
        }

        Assert.Equal("Value\n" + value.ToString(culture) + "\n", writer.ToString());
    }

    [Fact]
    public void Default_Writer_Quotes_Custom_Span_Formatted_Values_When_Needed()
    {
        using var writer = new StringWriter();
        using (var csv = new CsvObjectWriter(writer, new CsvSaveOptions { NewLine = "\n" }))
        {
            csv.WriteRow(new[] { "Value" }, new object?[] { new CommaSpanValue("A,B") });
        }

        Assert.Equal("Value\n\"A,B\"\n", writer.ToString());
    }

    private readonly struct CommaSpanValue : ISpanFormattable
    {
        private readonly string _value;

        public CommaSpanValue(string value)
        {
            _value = value;
        }

        public bool TryFormat(Span<char> destination, out int charsWritten, ReadOnlySpan<char> format, IFormatProvider? provider)
        {
            if (_value.AsSpan().TryCopyTo(destination))
            {
                charsWritten = _value.Length;
                return true;
            }

            charsWritten = 0;
            return false;
        }

        public string ToString(string? format, IFormatProvider? formatProvider) => _value;

        public override string ToString() => _value;
    }
#endif
}
