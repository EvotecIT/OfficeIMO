using System;
using System.IO;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvObjectWriterBatchTests
{
    [Fact]
    public void WriteRows_WritesSharedSchemaAndObjectValues()
    {
        object?[][] rows =
        [
            ["A", 1],
            ["B", 2]
        ];

        using var writer = new StringWriter();
        using (var csv = new CsvObjectWriter(writer, new CsvSaveOptions { NewLine = "\n" }, leaveOpen: true))
        {
            csv.WriteRows(["Name", "Value"], rows);
        }

        Assert.Equal("Name,Value\nA,1\nB,2\n", writer.ToString());
    }

    [Fact]
    public void WriteTextRows_EscapesEveryPreparedValue()
    {
        string?[][] rows =
        [
            ["A, quoted", "A\"B"],
            [null, "line 1\nline 2"]
        ];

        using var writer = new StringWriter();
        using (var csv = new CsvObjectWriter(writer, new CsvSaveOptions { NewLine = "\n" }, leaveOpen: true))
        {
            csv.WriteTextRows(["Name", "Notes"], rows);
        }

        Assert.Equal("Name,Notes\n\"A, quoted\",\"A\"\"B\"\n,\"line 1\nline 2\"\n", writer.ToString());
    }

    [Fact]
    public void WriteRows_RejectsRowsThatDoNotMatchTheSharedSchema()
    {
        object?[][] rows =
        [
            ["A", 1],
            ["B"]
        ];

        using var writer = new StringWriter();
        using var csv = new CsvObjectWriter(writer, new CsvSaveOptions { NewLine = "\n" }, leaveOpen: true);

        Assert.Throws<CsvException>(() => csv.WriteRows(["Name", "Value"], rows));
        Assert.Equal("Name,Value\nA,1\n", writer.ToString());
    }

    [Fact]
    public void WriteRows_WritesHeaderForAnEmptyBatch()
    {
        using var writer = new StringWriter();
        using (var csv = new CsvObjectWriter(writer, new CsvSaveOptions { NewLine = "\n" }, leaveOpen: true))
        {
            csv.WriteRows(["Name", "Value"], Array.Empty<object?[]>());
        }

        Assert.Equal("Name,Value\n", writer.ToString());
    }
}
