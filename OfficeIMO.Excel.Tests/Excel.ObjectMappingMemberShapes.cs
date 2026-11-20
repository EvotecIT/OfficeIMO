using System.Data;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public class ExcelObjectMappingMemberShapesTests {
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
    public void WriteDataSkipsIndexersAndWritesTheMostDerivedHiddenProperty() {
        string file = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        try {
            using (var document = ExcelDocument.Create(file)) {
                document.AddSheet("Indexed").WriteData(new[] { new IndexedRow { Name = "a", Count = 1 } });
                document.AddSheet("Hidden").WriteData(new[] { new HiddenDerivedRow { Value = "text", Name = "n" } });
                document.Save();
            }

            using var reader = ExcelDocumentReader.Open(file);
            object?[,] indexed = reader.GetSheet("Indexed").ReadRange("A1:C2");
            Assert.Equal("Name", indexed[0, 0]);
            Assert.Equal("Count", indexed[0, 1]);
            Assert.Null(indexed[0, 2]);

            object?[,] hidden = reader.GetSheet("Hidden").ReadRange("A1:C2");
            var headers = new[] { hidden[0, 0], hidden[0, 1], hidden[0, 2] };
            Assert.Equal(1, headers.Count(header => Equals(header, "Value")));
            Assert.Equal("text", hidden[1, Array.IndexOf(headers, "Value")]);
        } finally {
            File.Delete(file);
        }
    }

    [Fact]
    public void ObjectDataTableBuilderUsesTheMostDerivedHiddenProperty() {
        DataTable table = ExcelObjectDataTableBuilder.FromObjects(new object?[] { new HiddenDerivedRow { Value = "text", Name = "n" } });

        Assert.Equal(2, table.Columns.Count);
        Assert.Equal("text", table.Rows[0]["Value"]);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ReadObjectsBindTheMostDerivedHiddenPropertyAndSkipIndexers(bool stream) {
        string file = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        try {
            using (var document = ExcelDocument.Create(file)) {
                var sheet = document.AddSheet("Data");
                sheet.CellValue(1, 1, "Value");
                sheet.CellValue(1, 2, "Name");
                sheet.CellValue(1, 3, "Count");
                sheet.CellValue(2, 1, "text");
                sheet.CellValue(2, 2, "n");
                sheet.CellValue(2, 3, 5);
                document.Save();
            }

            using var reader = ExcelDocumentReader.Open(file);
            var data = reader.GetSheet("Data");
            HiddenDerivedRow hidden = Assert.Single(stream
                ? data.ReadObjectsStream<HiddenDerivedRow>("A1:C2")
                : data.ReadObjects<HiddenDerivedRow>("A1:C2"));
            IndexedRow indexed = Assert.Single(stream
                ? data.ReadObjectsStream<IndexedRow>("A1:C2")
                : data.ReadObjects<IndexedRow>("A1:C2"));

            Assert.Equal("text", hidden.Value);
            Assert.Equal("n", hidden.Name);
            Assert.Equal(5, indexed.Count);
        } finally {
            File.Delete(file);
        }
    }
}
