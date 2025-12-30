using System;
using System.Collections.Generic;
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
}
