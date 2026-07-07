using System.Linq;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvHeaderNormalizationTests
{
    [Fact]
    public void Duplicate_Headers_Are_Renamed_By_Default()
    {
        var parsed = CsvDocument.Parse("Name,Name,Value\nAlpha,Beta,1\n");
        var row = Assert.Single(parsed.AsEnumerable());

        Assert.Equal(new[] { "Name", "Name_2", "Value" }, parsed.Header);
        Assert.Equal("Alpha", row.AsString("Name"));
        Assert.Equal("Beta", row.AsString("Name_2"));
    }

    [Fact]
    public void Duplicate_Headers_Can_Be_Preserved()
    {
        var parsed = CsvDocument.Parse(
            "Name,Name\nAlpha,Beta\n",
            new CsvLoadOptions { DuplicateHeaderBehavior = CsvDuplicateHeaderBehavior.Preserve });

        Assert.Equal(new[] { "Name", "Name" }, parsed.Header);
        Assert.Equal("Alpha", parsed.AsEnumerable().Single().AsString("Name"));
    }

    [Fact]
    public void Duplicate_Headers_Can_Throw()
    {
        var ex = Assert.Throws<CsvException>(() => CsvDocument.Parse(
            "Name,Name\nAlpha,Beta\n",
            new CsvLoadOptions { DuplicateHeaderBehavior = CsvDuplicateHeaderBehavior.Throw }));

        Assert.Contains("duplicate column name 'Name'", ex.Message);
    }

    [Fact]
    public void Duplicate_Header_Renaming_Avoids_Existing_Names()
    {
        var parsed = CsvDocument.Parse("Name,Name,Name_2\nAlpha,Beta,Gamma\n");

        Assert.Equal(new[] { "Name", "Name_3", "Name_2" }, parsed.Header);
    }

    [Fact]
    public void W3C_Duplicate_Headers_Are_Normalized()
    {
        var parsed = CsvDocument.Parse("#Fields: date time time\n2026-07-07 10:00 10:01\n", new CsvLoadOptions { Delimiter = ' ' });
        var row = Assert.Single(parsed.AsEnumerable());

        Assert.Equal(new[] { "date", "time", "time_2" }, parsed.Header);
        Assert.Equal("10:00", row.AsString("time"));
        Assert.Equal("10:01", row.AsString("time_2"));
    }
}
