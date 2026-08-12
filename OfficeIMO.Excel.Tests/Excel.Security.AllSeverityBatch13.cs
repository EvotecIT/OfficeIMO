using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;
using OfficeIMO.Drawing;
using OfficeIMO.Data;
using System.Data;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelAllSeverityBatch13SecurityTests {
    [Fact]
    public async Task RemoteImageInsertionRejectsLoopbackDestinationByDefault() {
        using var stream = new MemoryStream();
        using ExcelDocument document = ExcelDocument.Create(stream);
        ExcelSheet sheet = document.AddWorksheet("Data");

        await Assert.ThrowsAsync<InvalidDataException>(() =>
            sheet.AddImageFromUrlAtAsync(1, 1, "http://127.0.0.1:6553/private.png"));
    }

    [Fact]
    public void SaveRejectsPackageMaterializationBeyondConfiguredLimit() {
        using var source = new MemoryStream();
        using ExcelDocument document = ExcelDocument.Create(source);
        document.AddWorksheet("Data").CellValue(1, 1, new string('x', 1024));
        using var destination = new MemoryStream();

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            document.Save(destination, new ExcelSaveOptions {
                DisableFastPackageWriter = true,
                MaxInMemoryPackageBytes = 64
            }));

        Assert.Contains("in-memory save limit", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void UnchangedSaveFastPathHonorsConfiguredMaterializationLimit() {
        byte[] package;
        using (var source = new MemoryStream())
        using (ExcelDocument created = ExcelDocument.Create(source)) {
            created.AddWorksheet("Data").CellValue(1, 1, "value");
            package = created.ToBytes();
        }

        using ExcelDocument document = ExcelDocument.Load(new MemoryStream(package, writable: false));
        using var destination = new MemoryStream();

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            document.Save(destination, new ExcelSaveOptions { MaxInMemoryPackageBytes = 64 }));

        Assert.Contains("in-memory save limit", exception.Message, StringComparison.Ordinal);
        Assert.Equal(0, destination.Length);
    }

    [Fact]
    public void SimpleFastSaveChecksLimitBeforeGrowingToBytesMemoryStream() {
        using var backing = new MemoryStream();
        using ExcelDocument document = ExcelDocument.Create(backing);
        document.AddWorksheet("Data").CellValue(1, 1, new string('x', 1024));

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            document.ToBytes(options: new ExcelSaveOptions { MaxInMemoryPackageBytes = 64 }));

        Assert.Contains("in-memory save limit", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DirectDataSetFastSaveChecksLimitBeforeGrowingToBytesMemoryStream() {
        var table = new DataTable("Data");
        table.Columns.Add("Value", typeof(string));
        table.Rows.Add(new string('x', 1024));
        var dataSet = new DataSet("Export");
        dataSet.Tables.Add(table);

        using var backing = new MemoryStream();
        using ExcelDocument document = ExcelDocument.Create(backing);
        document.InsertDataSet(dataSet, autoFit: false);

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            document.ToBytes(options: new ExcelSaveOptions { MaxInMemoryPackageBytes = 64 }));

        Assert.Contains("in-memory save limit", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DirectDataSetFastSaveRollsBackEmptyMemoryDestinationWhenLimitIsExceeded() {
        var table = new DataTable("Data");
        table.Columns.Add("Value", typeof(string));
        table.Rows.Add(new string('x', 1024));
        var dataSet = new DataSet("Export");
        dataSet.Tables.Add(table);

        using var backing = new MemoryStream();
        using ExcelDocument document = ExcelDocument.Create(backing);
        document.InsertDataSet(dataSet, autoFit: false);
        using var destination = new MemoryStream();

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            document.Save(destination, new ExcelSaveOptions { MaxInMemoryPackageBytes = 64 }));

        Assert.Contains("in-memory save limit", exception.Message, StringComparison.Ordinal);
        Assert.Equal(0, destination.Length);
        Assert.Equal(0, destination.Position);
    }

    [Fact]
    public void DirectDataSetFastSavePreservesPopulatedMemoryDestinationWhenLimitIsExceeded() {
        var table = new DataTable("Data");
        table.Columns.Add("Value", typeof(string));
        table.Rows.Add(new string('x', 1024));
        var dataSet = new DataSet("Export");
        dataSet.Tables.Add(table);

        using var backing = new MemoryStream();
        using ExcelDocument document = ExcelDocument.Create(backing);
        document.InsertDataSet(dataSet, autoFit: false);
        byte[] original = [1, 2, 3, 4];
        using var destination = new MemoryStream();
        destination.Write(original, 0, original.Length);
        destination.Position = 2;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            document.Save(destination, new ExcelSaveOptions { MaxInMemoryPackageBytes = 64 }));

        Assert.Contains("in-memory save limit", exception.Message, StringComparison.Ordinal);
        Assert.Equal(2, destination.Position);
        Assert.Equal(original, destination.ToArray());
    }

    [Fact]
    public void RowsFromStopsUnboundedSourceEnumerationAtConfiguredLimit() {
        using var stream = new MemoryStream();
        using ExcelDocument document = ExcelDocument.Create(stream);

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            document.AsFluent().Sheet("Data", sheet =>
                sheet.RowsFrom(InfiniteRows(), options => options.MaxRows = 3)));

        Assert.Contains("3-row", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowsFromStopsStreamingSourceAtCellLimit() {
        using var stream = new MemoryStream();
        using ExcelDocument document = ExcelDocument.Create(stream);
        int enumeratedRows = 0;

        IEnumerable<WideSimpleRow> StreamingRows() {
            while (true) {
                enumeratedRows++;
                yield return new WideSimpleRow { Left = enumeratedRows, Right = enumeratedRows };
            }
        }

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            document.AsFluent().Sheet("Data", sheet =>
                sheet.RowsFrom(StreamingRows(), options => {
                    options.MaxRows = 5;
                    options.MaxCells = 5;
                })));

        Assert.Contains("5-cell", exception.Message, StringComparison.Ordinal);
        Assert.Equal(2, enumeratedRows);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void RowsFromEmptyInputSkipsUnrenderedProjectionLimits(bool constrainColumns) {
        using var stream = new MemoryStream();
        using ExcelDocument document = ExcelDocument.Create(stream);

        Exception? exception = Record.Exception(() =>
            document.AsFluent().Sheet("Data", sheet =>
                sheet.RowsFrom(Array.Empty<WideSimpleRow>(), options => {
                    if (constrainColumns) {
                        options.MaxColumns = 1;
                    } else {
                        options.MaxCells = 1;
                    }
                })));

        Assert.Null(exception);
    }

    [Fact]
    public void TableFromStopsUnboundedSourceEnumerationAtConfiguredLimit() {
        using var stream = new MemoryStream();
        using ExcelDocument document = ExcelDocument.Create(stream);

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => document.Compose("Data", sheet => sheet.TableFrom(
                InfiniteRows(), configure: options => options.MaxRows = 3)));

        Assert.Contains("3-row", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowsFromStopsNestedExpansionAtConfiguredLimit() {
        using var stream = new MemoryStream();
        using ExcelDocument document = ExcelDocument.Create(stream);
        var rows = new[] { new NestedRow { Name = "item", Values = new[] { 1, 2, 3 } } };

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            document.AsFluent().Sheet("Data", sheet =>
                sheet.RowsFrom(rows, options => {
                    options.MaxRows = 2;
                    options.CollectionMode = CollectionMode.ExpandRows;
                })));

        Assert.Contains("nested expansion", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowsFromHonorsPerCollectionLimitDuringNestedExpansion() {
        using var stream = new MemoryStream();
        using ExcelDocument document = ExcelDocument.Create(stream);
        var rows = new[] { new NestedEnumerableRow { Values = InfiniteValues() } };

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            document.AsFluent().Sheet("Data", sheet =>
                sheet.RowsFrom(rows, options => {
                    options.MaxRows = 100;
                    options.MaxCollectionItems = 2;
                    options.CollectionMode = CollectionMode.ExpandRows;
                })));

        Assert.Contains("2-item", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowsFromSimpleFastPathRejectsColumnsBeyondConfiguredLimit() {
        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            SheetBuilder.EnsureRowsFromColumnLimit(2, 1));

        Assert.Contains("1-column", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowsFromHonorsFinalCellLimitDuringNestedExpansion() {
        using var stream = new MemoryStream();
        using ExcelDocument document = ExcelDocument.Create(stream);
        var rows = new[] { new NestedRow { Name = "A", Values = new[] { 1, 2, 3 } } };

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            document.AsFluent().Sheet("Data", sheet =>
                sheet.RowsFrom(rows, options => {
                    options.MaxCells = 5;
                    options.CollectionMode = CollectionMode.ExpandRows;
                })));

        Assert.Contains("5-cell", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ObjectFlattenerStopsUnboundedNestedCollectionEnumeration() {
        var flattener = new ObjectFlattener();
        var options = new ObjectFlattenerOptions { MaxCollectionItems = 2 };

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            flattener.Flatten(new NestedEnumerableRow { Values = InfiniteValues() }, options));

        Assert.Contains("2-item", exception.Message, StringComparison.Ordinal);
    }

    private static IEnumerable<SimpleRow> InfiniteRows() {
        int index = 0;
        while (true) yield return new SimpleRow { Value = index++ };
    }

    private static IEnumerable<int> InfiniteValues() {
        int index = 0;
        while (true) yield return index++;
    }

    private sealed class SimpleRow {
        public int Value { get; set; }
    }

    private sealed class WideSimpleRow {
        public int Left { get; set; }
        public int Right { get; set; }
    }

    private sealed class NestedRow {
        public string Name { get; set; } = string.Empty;
        public int[] Values { get; set; } = Array.Empty<int>();
    }

    private sealed class NestedEnumerableRow {
        public IEnumerable<int> Values { get; set; } = Array.Empty<int>();
    }
}
