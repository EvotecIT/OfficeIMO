using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
