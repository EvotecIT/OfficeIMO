using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelAllSeverityBatch16SecurityTests {
    [Fact]
    public void InsertObjectsRejectsCyclicDictionaryGraphs() {
        using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = document.AddWorksheet("Data");
        var row = new Dictionary<string, object?>();
        row["Self"] = row;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            sheet.InsertObjects(new[] { row }));

        Assert.Contains("reference cycle", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void InsertObjectsStopsSourcesAtRemainingWorksheetRows() {
        using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = document.AddWorksheet("Data");

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            sheet.InsertObjects(TwoRows(), includeHeaders: false, startRow: 1_048_576));

        Assert.Contains("1-row worksheet limit", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void InsertObjectsDoesNotTrustReadOnlyListCount() {
        using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = document.AddWorksheet("Data");

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            sheet.InsertObjects(new MisreportedReadOnlyList(), includeHeaders: false, startRow: 1_048_576));

        Assert.Contains("1-row worksheet limit", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void InsertObjectsRejectsDictionaryHeaderExpansionPastWorksheetColumns() {
        using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = document.AddWorksheet("Data");

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            sheet.InsertObjects(new object[] { new ExpandingReadOnlyDictionary() }, includeHeaders: false));

        Assert.Contains("16384-column limit", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void InsertObjectsRejectsDenseDictionaryProjectionPastCellBudget() {
        using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = document.AddWorksheet("Data");
        Dictionary<string, object?> row = Enumerable.Range(0, 1_024)
            .ToDictionary(index => $"Column{index}", index => (object?)index);

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            sheet.InsertObjects(Enumerable.Repeat<object>(row, 4_096)));

        Assert.Contains("cell safety limit", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    private static IEnumerable<object> TwoRows() {
        yield return new { Value = 1 };
        yield return new { Value = 2 };
    }

    private sealed class MisreportedReadOnlyList : IReadOnlyList<object> {
        public object this[int index] => new { Value = index };

        public int Count => 1;

        public IEnumerator<object> GetEnumerator() {
            while (true) {
                yield return new { Value = 1 };
            }
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }

    private sealed class ExpandingReadOnlyDictionary : IReadOnlyDictionary<string, object?> {
        public object? this[string key] => null;

        public IEnumerable<string> Keys => Enumerable.Empty<string>();

        public IEnumerable<object?> Values => Enumerable.Empty<object?>();

        public int Count => 1;

        public bool ContainsKey(string key) => false;

        public bool TryGetValue(string key, out object? value) {
            value = null;
            return false;
        }

        public IEnumerator<KeyValuePair<string, object?>> GetEnumerator() {
            for (int index = 0; index <= 16_384; index++) {
                yield return new KeyValuePair<string, object?>($"Column{index}", index);
            }
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
