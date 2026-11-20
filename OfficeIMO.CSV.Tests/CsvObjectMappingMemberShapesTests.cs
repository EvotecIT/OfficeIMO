using OfficeIMO.CSV;
using OfficeIMO.Data;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvObjectMappingMemberShapesTests {
    public class IndexedRow {
        private readonly string[] _slots = { "x", "y" };
        public string Name { get; set; } = string.Empty;
        public int Count { get; set; }
        public string this[int index] { get => _slots[index]; set => _slots[index] = value; }
    }

    public class HiddenBaseRow {
        public int Value { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    public class HiddenDerivedRow : HiddenBaseRow {
        public new string Value { get; set; } = string.Empty;
    }

    [Fact]
    public void ObjectWritersSkipIndexersAndUseTheMostDerivedHiddenProperty() {
        CsvDocument indexed = CsvDocument.FromObjects(new object?[] { new IndexedRow { Name = "a", Count = 1 } });
        Assert.Equal(new[] { "Name", "Count" }, indexed.Header);

        CsvDocument hidden = CsvDocument.FromObjects(new object?[] { new HiddenDerivedRow { Value = "text", Name = "n" } });
        Assert.Equal(1, hidden.Header.Count(header => header == "Value"));
        Assert.Equal("text", hidden.AsEnumerable().Single().AsString("Value"));

        using var writer = new StringWriter();
        CsvDocument.WriteObjects(writer, new object?[] { new HiddenDerivedRow { Value = "text", Name = "n" } },
            new CsvSaveOptions { NewLine = "\n" });
        string[] lines = writer.ToString().Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
        string[] header = lines[0].Split(',');
        Assert.Equal(1, header.Count(column => column == "Value"));
        Assert.Equal("text", lines[1].Split(',')[Array.IndexOf(header, "Value")]);
    }

    [Fact]
    public void TypedReadsSkipIndexersAndBindTheMostDerivedHiddenProperty() {
        CsvDocument document = CsvDocument.Parse("Value,Name,Count\ntext,n,5\n");

        HiddenDerivedRow hidden = Assert.Single(document.RowsAs<HiddenDerivedRow>());
        Assert.Equal("text", hidden.Value);
        Assert.Equal("n", hidden.Name);

        IndexedRow indexed = Assert.Single(document.RowsAs<IndexedRow>());
        Assert.Equal(5, indexed.Count);
    }
}
