using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderTableProjectionIdentityTests {
    [Theory]
    [InlineData("SourceBlockIndex")]
    [InlineData("BlockIndex")]
    [InlineData("StartLine")]
    [InlineData("A1Range")]
    [InlineData("HeadingPath")]
    [InlineData("HeadingSlug")]
    [InlineData("SourceBlockKind")]
    [InlineData("Path")]
    [InlineData("TableIndex")]
    public void EqualPayloadsWithConflictingSourcePositionsRemainSeparate(string locationField) {
        var first = new ReaderLocation { Path = "doc" };
        var second = new ReaderLocation { Path = "doc", Page = 1 };
        var property = typeof(ReaderLocation).GetProperty(locationField)!;
        property.SetValue(first, property.PropertyType == typeof(string) ? "second-position" : 2);
        property.SetValue(second, property.PropertyType == typeof(string) ? "first-position" : 1);
        var aggregate = new ReaderTable { Columns = new[] { "Item" }, Rows = new[] { new[] { "Widget" } }, Location = first };
        var pageOnly = new ReaderTable { Columns = aggregate.Columns, Rows = aggregate.Rows, Location = second };
        var source = new OfficeDocumentReadResult { Tables = new[] { aggregate },
            Pages = new[] { new OfficeDocumentPage { Number = 1, Tables = new[] { pageOnly } } } };
        foreach (var document in new[] { source, OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(source)) }) {
            ReaderTable[] tables = document.EnumerateTables().ToArray();
            Assert.Equal(2, tables.Length);
            Assert.Null(tables[0].Location!.Page);
            Assert.Equal(1, tables[1].Location!.Page);
            Assert.Equal(property.GetValue(first), property.GetValue(tables[0].Location));
            Assert.Equal(property.GetValue(second), property.GetValue(tables[1].Location));
        }
        Assert.Null(first.Page);
    }
}
