using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvDocumentFromObjectsTests
{
    [Fact]
    public void FromObjects_UsesPropertiesAndDelimiter()
    {
        var items = new[]
        {
            new { Name = "A", Value = 1 },
            new { Name = "B", Value = 2 }
        };

        var doc = CsvDocument.FromObjects(items, ';');

        Assert.Equal(';', doc.Delimiter);
        Assert.Equal(new[] { "Name", "Value" }, doc.Header);

        var rows = doc.AsEnumerable().ToList();
        Assert.Equal(2, rows.Count);
        Assert.Equal("A", rows[0].AsString("Name"));
        Assert.Equal(1, rows[0].AsInt32("Value"));
        Assert.Equal("B", rows[1].AsString("Name"));
        Assert.Equal(2, rows[1].AsInt32("Value"));
    }

    [Fact]
    public void FromObjects_UsesDictionaryKeys()
    {
        var items = new List<Dictionary<string, object?>>
        {
            new() { ["Name"] = "C", ["Value"] = 3 },
            new() { ["Name"] = "D", ["Value"] = 4 }
        };

        var doc = CsvDocument.FromObjects(items);

        Assert.Contains("Name", doc.Header);
        Assert.Contains("Value", doc.Header);
        Assert.Equal(2, doc.AsEnumerable().Count());
        Assert.Equal("C", doc.AsEnumerable().ElementAt(0).AsString("Name"));    
        Assert.Equal(4, doc.AsEnumerable().ElementAt(1).AsInt32("Value"));      
    }

    [Fact]
    public void WriteObjects_WritesWithoutMaterializingDocument()
    {
        var items = new object?[]
        {
            new { Name = "A", Value = 1 },
            new { Name = "B, quoted", Value = 2 }
        };

        using var writer = new StringWriter();
        CsvDocument.WriteObjects(writer, items, new CsvSaveOptions { NewLine = "\n" });

        Assert.Equal("Name,Value\nA,1\n\"B, quoted\",2\n", writer.ToString());
    }

    [Fact]
    public void CsvObjectWriter_StreamsRowsAndCanLeaveWriterOpen()
    {
        using var writer = new StringWriter();
        using (var csvWriter = new CsvObjectWriter(writer, new CsvSaveOptions { NewLine = "\n" }, leaveOpen: true))
        {
            csvWriter.WriteObject(new { Name = "A", Value = 1 });
            csvWriter.WriteObject(new { Name = "B", Value = 2 });
            Assert.True(csvWriter.HasRows);
        }

        writer.Write("#");
        Assert.Equal("Name,Value\nA,1\nB,2\n#", writer.ToString());
    }

    [Fact]
    public void CsvObjectWriter_WritesProjectedRows()
    {
        using var writer = new StringWriter();
        using (var csvWriter = new CsvObjectWriter(writer, new CsvSaveOptions { NewLine = "\n" }, leaveOpen: true))
        {
            csvWriter.WriteRow(new[] { "Name", "Value" }, new object?[] { "A", 1 });
            csvWriter.WriteRow(new[] { "Name", "Value" }, new object?[] { "B", 2 });
        }

        Assert.Equal("Name,Value\nA,1\nB,2\n", writer.ToString());
    }

    [Fact]
    public void ReadRows_StreamsHeaderAndRows()
    {
        using var reader = new StringReader("Name,Value\nA,1\nB,2\n");
        var rows = new List<string>();

        CsvDocument.ReadRows(reader, (header, values) =>
        {
            Assert.Equal(new[] { "Name", "Value" }, header);
            rows.Add(values[0] + ":" + values[1]);
        });

        Assert.Equal(new[] { "A:1", "B:2" }, rows);
    }

    [Fact]
    public void ReadRowsReusable_StreamsHeaderAndRows()
    {
        using var reader = new StringReader("Name,Value\nA,1\nB,2\n");
        var rows = new List<string>();

        CsvDocument.ReadRowsReusable(reader, (header, values) =>
        {
            Assert.Equal(new[] { "Name", "Value" }, header);
            rows.Add(values[0] + ":" + values[1]);
        });

        Assert.Equal(new[] { "A:1", "B:2" }, rows);
    }

    [Fact]
    public void ReadRowsReusable_ReusesUnquotedRowBuffer()
    {
        using var reader = new StringReader("Name,Value\nA,1\nB,2\n");
        IReadOnlyList<string>? first = null;
        IReadOnlyList<string>? second = null;

        CsvDocument.ReadRowsReusable(reader, (_, values) =>
        {
            if (first == null)
            {
                first = values;
            }
            else
            {
                second = values;
            }
        });

        Assert.NotNull(first);
        Assert.Same(first, second);
    }

    [Fact]
    public void ReadRowsReusable_HandlesQuotedRows()
    {
        using var reader = new StringReader("Name,Value\n\"A, quoted\",1\nB,2\n");
        var rows = new List<string>();

        CsvDocument.ReadRowsReusable(reader, (_, values) =>
        {
            rows.Add(values[0] + ":" + values[1]);
        });

        Assert.Equal(new[] { "A, quoted:1", "B:2" }, rows);
    }

    [Fact]
    public void ReadRecordsReusable_StreamsRawHeaderAndRows()
    {
        using var reader = new StringReader("Name,Value\nA,1\nB,2\n");
        var records = new List<string>();

        CsvDocument.ReadRecordsReusable(reader, values =>
        {
            records.Add(values[0] + ":" + values[1]);
        });

        Assert.Equal(new[] { "Name:Value", "A:1", "B:2" }, records);
    }

    [Fact]
    public void ReadRecordsReusable_PreservesExtraFields()
    {
        using var reader = new StringReader("Name\nA,1\n");
        IReadOnlyList<string>? row = null;

        CsvDocument.ReadRecordsReusable(reader, values =>
        {
            if (values.Count > 1)
            {
                row = values.ToArray();
            }
        });

        Assert.NotNull(row);
        Assert.Equal(new[] { "A", "1" }, row);
    }

    [Fact]
    public void SaveObjects_WritesFile()
    {
        var path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.SaveObjects." + Guid.NewGuid().ToString("N") + ".csv");
        try
        {
            CsvDocument.SaveObjects(path, new object?[] { new { Name = "A", Value = 1 } }, new CsvSaveOptions { NewLine = "\n" });

            Assert.Equal("Name,Value\nA,1\n", File.ReadAllText(path));
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
    public void FromObjects_ThrowsOnNullItems()
    {
        var ex = Assert.Throws<ArgumentNullException>(() => CsvDocument.FromObjects(null!));
        Assert.Equal("items", ex.ParamName);
    }

    [Fact]
    public void FromObjects_ThrowsOnEmptyItems()
    {
        var ex = Assert.Throws<ArgumentException>(() => CsvDocument.FromObjects(Array.Empty<object?>()));
        Assert.StartsWith("Provide at least one data row.", ex.Message);
        Assert.Equal("items", ex.ParamName);
    }

    [Fact]
    public void FromObjects_ThrowsOnNullFirstRow()
    {
        var ex = Assert.Throws<ArgumentException>(() => CsvDocument.FromObjects(new object?[] { null }));
        Assert.StartsWith("Data rows cannot be null.", ex.Message);
        Assert.Equal("items", ex.ParamName);
    }

    [Fact]
    public void FromObjects_ThrowsOnNullEntry()
    {
        var ex = Assert.Throws<InvalidOperationException>(() => CsvDocument.FromObjects(new object?[] { new { Name = "A", Value = 1 }, null }));
        Assert.Equal("Data rows cannot contain null entries.", ex.Message);
    }

    [Fact]
    public void FromObjects_ThrowsOnMissingColumns()
    {
        var ex = Assert.Throws<InvalidOperationException>(() => CsvDocument.FromObjects(new object?[] { new object() }));
        Assert.Equal("Unable to infer column names. Use objects with properties or dictionaries.", ex.Message);
    }
}
