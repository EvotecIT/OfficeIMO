using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvLoadFeatureParityTests
{
    [Fact]
    public void Load_Appends_Static_Columns_To_Materialized_Rows()
    {
        var parsed = CsvDocument.Parse(
            "Name,Value\nAlpha,1\nBeta,2\n",
            new CsvLoadOptions {
                StaticColumns = new Dictionary<string, object?> {
                    ["SourceFile"] = "input.csv",
                    ["BatchId"] = 42
                }
            });

        Assert.Equal(new[] { "Name", "Value", "SourceFile", "BatchId" }, parsed.Header);
        var row = parsed.AsEnumerable().First();
        Assert.Equal("input.csv", row.AsString("SourceFile"));
        Assert.Equal(42, row.Get<int>("BatchId"));
    }

    [Fact]
    public void Load_Appends_Static_Columns_In_Streaming_Mode()
    {
        var parsed = CsvDocument.Parse(
            "Name,Value\nAlpha,1\n",
            new CsvLoadOptions {
                Mode = CsvLoadMode.Stream,
                StaticColumns = new Dictionary<string, object?> {
                    ["SourceFile"] = "stream.csv"
                }
            });

        var row = Assert.Single(parsed.AsEnumerable());

        Assert.Equal(new[] { "Name", "Value", "SourceFile" }, parsed.Header);
        Assert.Equal("stream.csv", row.AsString("SourceFile"));
    }

    [Fact]
    public void Load_Converts_Configured_Null_Token_To_Null()
    {
        var parsed = CsvDocument.Parse(
            "Name,Value\nAlpha,<null>\n",
            new CsvLoadOptions { NullValue = "<null>" });

        var row = Assert.Single(parsed.AsEnumerable());

        Assert.Null(row["Value"]);
    }

    [Fact]
    public void Rows_Use_Custom_DateTime_Formats_For_Typed_Conversion()
    {
        var parsed = CsvDocument.Parse(
            "Created\n07-Jul-2026\n",
            new CsvLoadOptions { DateTimeFormats = new[] { "dd-MMM-yyyy" } });

        var row = Assert.Single(parsed.AsEnumerable());

        Assert.Equal(new DateTime(2026, 7, 7), row.AsDateTime("Created"));
    }

    [Fact]
    public void Quote_Parsing_Defaults_To_Lenient_Mode()
    {
        var parsed = CsvDocument.Parse("Name,Value\nAlpha,\"one\"two\n");

        var row = Assert.Single(parsed.AsEnumerable());

        Assert.Equal("onetwo", row.AsString("Value"));
    }

    [Fact]
    public void Quote_Parsing_Strict_Mode_Rejects_Invalid_Quoted_Field()
    {
        var ex = Assert.Throws<CsvParseException>(() =>
            CsvDocument.Parse(
                "Name,Value\nAlpha,\"one\"two\n",
                new CsvLoadOptions { QuoteParsingMode = CsvQuoteParsingMode.Strict }));

        Assert.Contains("Invalid quoted field", ex.Message);
    }

    [Fact]
    public void Schema_Validation_Uses_Custom_DateTime_Formats()
    {
        var parsed = CsvDocument.Parse(
            "Created\n07-Jul-2026\n",
            new CsvLoadOptions { DateTimeFormats = new[] { "dd-MMM-yyyy" } })
            .EnsureSchema(schema => schema.Column("Created").AsDateTime().Required());

        parsed.Validate(out var errors);

        Assert.Empty(errors);
    }

    [Fact]
    public void ReadRows_Appends_Static_Columns()
    {
        var seen = new List<string>();

        CsvDocument.ReadRows(
            new StringReader("Name,Value\nAlpha,1\n"),
            (header, values) => seen.Add($"{string.Join("|", header)}={string.Join("|", values)}"),
            new CsvLoadOptions {
                StaticColumns = new Dictionary<string, object?> {
                    ["SourceFile"] = "read.csv"
                }
            });

        Assert.Equal("Name|Value|SourceFile=Alpha|1|read.csv", Assert.Single(seen));
    }

    [Fact]
    public void Save_Uses_Null_Token_And_DateTime_Format()
    {
        var document = new CsvDocument()
            .WithHeader("Name", "Created", "Value")
            .AddRow("Alpha", new DateTime(2026, 7, 7, 13, 45, 0, DateTimeKind.Utc), null);

        var text = document.ToString(new CsvSaveOptions {
            NewLine = "\n",
            NullValue = "<null>",
            DateTimeFormat = "yyyyMMdd-HHmm"
        });

        Assert.Equal("Name,Created,Value\nAlpha,20260707-1345,<null>\n", text);
    }

    [Fact]
    public void WriteObjects_Uses_Null_Token_And_DateTime_Format()
    {
        var rows = new[] {
            new {
                Name = "Alpha",
                Created = new DateTime(2026, 7, 7, 13, 45, 0, DateTimeKind.Utc),
                Value = (string?)null
            }
        };

        using var writer = new StringWriter();
        CsvDocument.WriteObjects(
            writer,
            rows,
            new CsvSaveOptions {
                NewLine = "\n",
                NullValue = "<null>",
                DateTimeFormat = "yyyyMMdd-HHmm"
            });

        Assert.Equal("Name,Created,Value\nAlpha,20260707-1345,<null>\n", writer.ToString());
    }

    [Fact]
    public void CsvObjectWriter_TextRows_Use_Null_Token_When_Configured()
    {
        using var writer = new StringWriter();
        using (var csv = new CsvObjectWriter(
            writer,
            new CsvSaveOptions { NewLine = "\n", NullValue = "<null>" },
            leaveOpen: true))
        {
            csv.WriteTextRow(new[] { "Name", "Value" }, new string?[] { "Alpha", null });
        }

        Assert.Equal("Name,Value\nAlpha,<null>\n", writer.ToString());
    }

    [Fact]
    public void Load_Honors_Cancellation_Token()
    {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            CsvDocument.Parse(
                "Name\nAlpha\n",
                new CsvLoadOptions { CancellationToken = cancellation.Token }));
    }

    [Fact]
    public void Load_Reports_Progress_At_Configured_Interval()
    {
        var reports = new List<long>();

        CsvDocument.Parse(
            "Name\nAlpha\nBeta\nGamma\n",
            new CsvLoadOptions {
                ProgressReportInterval = 2,
                ProgressCallback = progress => reports.Add(progress.RecordsRead)
            });

        Assert.Equal(new[] { 2L, 4L }, reports);
    }

    [Fact]
    public void Load_Collects_And_Skips_Parse_Errors_When_Configured()
    {
        var errors = new List<CsvParseError>();
        var parsed = CsvDocument.Parse(
            "Name,Value\nAlpha,1\nBroken,\"one\"two\nBeta,2\n",
            new CsvLoadOptions {
                QuoteParsingMode = CsvQuoteParsingMode.Strict,
                ParseErrorAction = CsvParseErrorAction.SkipRow,
                CollectParseErrors = true,
                ParseErrors = errors
            });

        Assert.Equal(new[] { "Alpha", "Beta" }, parsed.AsEnumerable().Select(row => row.AsString("Name")).ToArray());
        var error = Assert.Single(errors);
        Assert.Contains("Invalid quoted field", error.Message);
    }

    [Fact]
    public void Load_Skips_Field_Limit_Parse_Errors_When_Configured()
    {
        var errors = new List<CsvParseError>();
        var parsed = CsvDocument.Parse(
            "Name\nOk\nTooLong\nFine\n",
            new CsvLoadOptions
            {
                MaxFieldLength = 4,
                ParseErrorAction = CsvParseErrorAction.SkipRow,
                CollectParseErrors = true,
                ParseErrors = errors
            });

        Assert.Equal(new[] { "Ok", "Fine" }, parsed.AsEnumerable().Select(row => row.AsString("Name")).ToArray());
        var error = Assert.Single(errors);
        Assert.Contains("exceeds the configured maximum", error.Message);
    }

    [Fact]
    public void Load_Rejects_Fields_Over_Configured_Limit()
    {
        var ex = Assert.Throws<CsvParseException>(() =>
            CsvDocument.Parse(
                "Name\nTooLong\n",
                new CsvLoadOptions { MaxFieldLength = 3 }));

        Assert.Contains("exceeds the configured maximum", ex.Message);
    }

    [Fact]
    public void Load_Normalizes_Smart_Quotes_When_Configured()
    {
        var parsed = CsvDocument.Parse(
            "Name,Value\nAlpha,\u201CHello\u201D\n",
            new CsvLoadOptions { NormalizeQuotes = true });

        Assert.Equal("\"Hello\"", Assert.Single(parsed.AsEnumerable()).AsString("Value"));
    }

    [Fact]
    public void Load_Reuses_Repeated_String_Instances_When_Configured()
    {
        var parsed = CsvDocument.Parse(
            "Name\nAlpha\nAlpha\n",
            new CsvLoadOptions { InternStrings = true });

        var rows = parsed.AsEnumerable().ToArray();
        Assert.Same(rows[0]["Name"], rows[1]["Name"]);
    }

    [Fact]
    public void Load_Reads_MultiCharacter_Delimiter()
    {
        var parsed = CsvDocument.Parse(
            "Name||Value\nAlpha||\"one||two\"\nBeta||2\n",
            new CsvLoadOptions { DelimiterText = "||" });

        var rows = parsed.AsEnumerable().ToArray();

        Assert.Equal(new[] { "Name", "Value" }, parsed.Header);
        Assert.Equal("one||two", rows[0].AsString("Value"));
        Assert.Equal("Beta", rows[1].AsString("Name"));
    }

    [Fact]
    public void Load_Reads_SingleCharacter_DelimiterText()
    {
        var parsed = CsvDocument.Parse(
            "Name\tValue\nAlpha\t1\n",
            new CsvLoadOptions { DelimiterText = "\t" });

        var row = Assert.Single(parsed.AsEnumerable());
        Assert.Equal(new[] { "Name", "Value" }, parsed.Header);
        Assert.Equal("Alpha", row.AsString("Name"));
        Assert.Equal("1", row.AsString("Value"));
    }

    [Fact]
    public void ReadRowsReusable_Reads_MultiCharacter_Delimiter()
    {
        var rows = new List<string>();

        CsvDocument.ReadRowsReusable(
            new StringReader("Name||Value\nAlpha||1\n"),
            (header, values) => rows.Add($"{string.Join(",", header)}={string.Join(",", values)}"),
            new CsvLoadOptions { DelimiterText = "||" });

        Assert.Equal("Name,Value=Alpha,1", Assert.Single(rows));
    }

    [Fact]
    public void ReadRowsReusable_Reads_SingleCharacter_DelimiterText()
    {
        var rows = new List<string>();

        CsvDocument.ReadRowsReusable(
            new StringReader("Name;Value\nAlpha;1\n"),
            (header, values) => rows.Add($"{string.Join(",", header)}={string.Join(",", values)}"),
            new CsvLoadOptions { DelimiterText = ";" });

        Assert.Equal("Name,Value=Alpha,1", Assert.Single(rows));
    }

#if NET8_0_OR_GREATER
    [Fact]
    public void ReadFieldSpans_Reads_MultiCharacter_Delimiter()
    {
        var fields = new List<string>();

        CsvDocument.ReadFieldSpansFromText(
            "Name||Value\nAlpha||\"one||two\"\n",
            (recordIndex, fieldIndex, value) => fields.Add($"{recordIndex}:{fieldIndex}:{value.ToString()}"),
            new CsvLoadOptions { DelimiterText = "||" });

        Assert.Equal(new[] { "0:0:Name", "0:1:Value", "1:0:Alpha", "1:1:one||two" }, fields);
    }

    [Fact]
    public void ReadFieldSpans_Reads_SingleCharacter_DelimiterText()
    {
        var fields = new List<string>();

        CsvDocument.ReadFieldSpansFromText(
            "Name\tValue\nAlpha\t1\n",
            (recordIndex, fieldIndex, value) => fields.Add($"{recordIndex}:{fieldIndex}:{value.ToString()}"),
            new CsvLoadOptions { DelimiterText = "\t" });

        Assert.Equal(new[] { "0:0:Name", "0:1:Value", "1:0:Alpha", "1:1:1" }, fields);
    }

    [Fact]
    public void ReadFieldSpans_Skips_Parse_Errors_When_Configured()
    {
        var fields = new List<string>();
        var errors = new List<CsvParseError>();

        CsvDocument.ReadFieldSpansFromText(
            "Name,Value\nAlpha,1\nBroken,\"one\"two\nBeta,2\n",
            (recordIndex, fieldIndex, value) => fields.Add($"{recordIndex}:{fieldIndex}:{value.ToString()}"),
            new CsvLoadOptions
            {
                QuoteParsingMode = CsvQuoteParsingMode.Strict,
                ParseErrorAction = CsvParseErrorAction.SkipRow,
                CollectParseErrors = true,
                ParseErrors = errors
            });

        Assert.Equal(new[] { "0:0:Name", "0:1:Value", "1:0:Alpha", "1:1:1", "2:0:Beta", "2:1:2" }, fields);
        var error = Assert.Single(errors);
        Assert.Contains("Invalid quoted field", error.Message);
    }
#endif

    [Fact]
    public void Save_Writes_MultiCharacter_Delimiter()
    {
        var document = new CsvDocument()
            .WithHeader("Name", "Value")
            .AddRow("Alpha", "one||two");

        var text = document.ToString(new CsvSaveOptions { DelimiterText = "||", NewLine = "\n" });

        Assert.Equal("Name||Value\nAlpha||\"one||two\"\n", text);
    }

    [Fact]
    public void Save_Writes_SingleCharacter_DelimiterText()
    {
        var document = new CsvDocument()
            .WithHeader("Name", "Value")
            .AddRow("Alpha", "1");

        var text = document.ToString(new CsvSaveOptions { DelimiterText = "\t", NewLine = "\n" });

        Assert.Equal("Name\tValue\nAlpha\t1\n", text);
    }

    [Fact]
    public void CsvObjectWriter_Writes_MultiCharacter_Delimiter()
    {
        using var writer = new StringWriter();
        using (var csv = new CsvObjectWriter(writer, new CsvSaveOptions { DelimiterText = "||", NewLine = "\n" }, leaveOpen: true))
        {
            csv.WriteObject(new { Name = "Alpha", Value = "one||two" });
        }

        Assert.Equal("Name||Value\nAlpha||\"one||two\"\n", writer.ToString());
    }

    [Fact]
    public void CsvObjectWriter_Writes_SingleCharacter_DelimiterText()
    {
        using var writer = new StringWriter();
        using (var csv = new CsvObjectWriter(writer, new CsvSaveOptions { DelimiterText = ";", NewLine = "\n" }, leaveOpen: true))
        {
            csv.WriteObject(new { Name = "Alpha", Value = 1 });
        }

        Assert.Equal("Name;Value\nAlpha;1\n", writer.ToString());
    }

    [Fact]
    public void WriteDataReader_Writes_MultiCharacter_Delimiter()
    {
        using var table = new DataTable();
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Value", typeof(string));
        table.Rows.Add("Alpha", "one||two");

        using var reader = table.CreateDataReader();
        using var writer = new StringWriter();
        CsvDocument.WriteDataReader(writer, reader, new CsvSaveOptions { DelimiterText = "||", NewLine = "\n" });

        Assert.Equal("Name||Value\nAlpha||\"one||two\"\n", writer.ToString());
    }

    [Fact]
    public void Save_NoClobber_Throws_When_File_Exists()
    {
        var path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.NoClobber." + Guid.NewGuid().ToString("N") + ".csv");
        try
        {
            File.WriteAllText(path, "existing");
            var document = new CsvDocument().WithHeader("Name").AddRow("Alpha");

            Assert.Throws<IOException>(() => document.Save(path, new CsvSaveOptions { NoClobber = true }));
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
    public void Save_Appends_To_Existing_File_When_Requested()
    {
        var path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.Append." + Guid.NewGuid().ToString("N") + ".csv");
        try
        {
            File.WriteAllText(path, "Name\nAlpha\n");
            new CsvDocument()
                .WithHeader("Name")
                .AddRow("Beta")
                .Save(path, new CsvSaveOptions { Append = true, IncludeHeader = false, NewLine = "\n" });

            Assert.Equal("Name\nAlpha\nBeta\n", File.ReadAllText(path));
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }
}
