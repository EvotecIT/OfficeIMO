using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderTableProjectionIdentityTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PromotedTablesAndVisualsPreserveHierarchyWithoutDuplicates(bool roundTrip) {
        var location = new ReaderLocation { Path = "doc.md", Page = 1, BlockAnchor = "section", HeadingPath = "A > B" };
        ReaderHeadingPath.SetHierarchyPath(location, ReaderHeadingPath.Combine(new[] { "A > B" }));
        var chunks = new[] { new ReaderChunk {
            Location = location,
            Tables = new[] { new ReaderTable { Columns = new[] { "Item" }, Rows = new[] { new[] { "Widget" } } } },
            Visuals = new[] { new ReaderVisual { Kind = "mermaid", Content = "graph TD\nA-->B", PayloadHash = "graph" } }
        } };
        var document = new OfficeDocumentReadResult {
            Chunks = chunks,
            Tables = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ExtractTables(chunks),
            Visuals = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ExtractVisuals(chunks)
        };
        if (roundTrip) document = OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(document));
        var table = Assert.Single(document.EnumerateTables());
        var visual = Assert.Single(OfficeDocumentModelTraversal.Visuals(document));
        Assert.Equal(location.HierarchyHeadingPath, table.Location!.HierarchyHeadingPath);
        Assert.Equal(location.HierarchyHeadingPath, visual.Location!.HierarchyHeadingPath);
    }

    [Theory]
    [InlineData("SourceBlockIndex")]
    [InlineData("BlockIndex")]
    [InlineData("StartLine")]
    [InlineData("EndLine")]
    [InlineData("NormalizedStartLine")]
    [InlineData("NormalizedEndLine")]
    [InlineData("A1Range")]
    [InlineData("HeadingPath")]
    [InlineData("HierarchyHeadingPath")]
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
