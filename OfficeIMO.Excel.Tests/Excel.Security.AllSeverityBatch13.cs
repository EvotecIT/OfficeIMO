using OfficeIMO.Excel;
using OfficeIMO.Drawing;
using OfficeIMO.Data;
using System.Data;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;
public sealed class ExcelAllSeverityBatch13SecurityTests
{
    [Fact]
    public async Task RemoteImageInsertionRejectsLoopbackDestinationByDefault() {
        using var stream = new MemoryStream();
        using ExcelDocument document = ExcelDocument.Create(stream);
        ExcelSheet sheet = document.AddSheet("Data");

        await Assert.ThrowsAsync<InvalidDataException>(() =>
            sheet.AddImageAsync(1, 1, new System.Uri("http://127.0.0.1:6553/private.png")));
    }

    [Fact]
    public void SaveRejectsPackageMaterializationBeyondConfiguredLimit() {
        using var source = new MemoryStream();
        using ExcelDocument document = ExcelDocument.Create(source);
        document.AddSheet("Data").CellValue(1, 1, new string('x', 1024));
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
            created.AddSheet("Data").CellValue(1, 1, "value");
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
        document.AddSheet("Data").CellValue(1, 1, new string('x', 1024));

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
    public void RowsFromStopsUnboundedSourceEnumerationAtConfiguredLimit()
    {
        using var stream = new MemoryStream();
        using ExcelDocument document = ExcelDocument.Create(stream);
        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
        {
            var sheet1 = document.AddSheet("Data");
            int nextRow2 = 1;
            var projection3 = new OfficeIMO.Data.ObjectFlattenerOptions();
            (projection3).MaxRows = 3;
            var table4 = sheet1.AddTable(InfiniteRows(), new ExcelTableOptions { Name = null, Style = ExcelTableStyle.TableStyleMedium2, IncludeAutoFilter = true, Projection = projection3 }, row: nextRow2);
            nextRow2 = table4.AsRange().LastRow + 1;
        });
        Assert.Contains("3-row", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowsFromStopsStreamingSourceAtCellLimit()
    {
        using var stream = new MemoryStream();
        using ExcelDocument document = ExcelDocument.Create(stream);
        int enumeratedRows = 0;
        IEnumerable<WideSimpleRow> StreamingRows()
        {
            while (true)
            {
                enumeratedRows++;
                yield return new WideSimpleRow
                {
                    Left = enumeratedRows,
                    Right = enumeratedRows
                };
            }
        }

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
        {
            var sheet5 = document.AddSheet("Data");
            int nextRow6 = 1;
            var projection7 = new OfficeIMO.Data.ObjectFlattenerOptions();
            {
                (projection7).MaxRows = 5;
                (projection7).MaxCells = 5;
            }

            var table8 = sheet5.AddTable(StreamingRows(), new ExcelTableOptions { Name = null, Style = ExcelTableStyle.TableStyleMedium2, IncludeAutoFilter = true, Projection = projection7 }, row: nextRow6);
            nextRow6 = table8.AsRange().LastRow + 1;
        });
        Assert.Contains("5-cell", exception.Message, StringComparison.Ordinal);
        Assert.Equal(2, enumeratedRows);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void EmptyNativeTableHeadersHonorProjectionLimits(bool constrainColumns)
    {
        using ExcelDocument document = ExcelDocument.Create();
        var sheet = document.AddSheet("Data");
        var projection = new ObjectFlattenerOptions();
        if (constrainColumns)
            projection.MaxColumns = 1;
        else
            projection.MaxCells = 1;
        Assert.Throws<InvalidDataException>(() => sheet.AddTable(Array.Empty<WideSimpleRow>(), new ExcelTableOptions { Projection = projection }));
        Assert.Empty(document.Inspect().Worksheets.Single().Cells);
    }

    [Fact]
    public void TableFromStopsUnboundedSourceEnumerationAtConfiguredLimit()
    {
        using var stream = new MemoryStream();
        using ExcelDocument document = ExcelDocument.Create(stream);
        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => ((System.Action<OfficeIMO.Excel.ExcelReportBuilder>)(sheet => sheet.AddTable(InfiniteRows(), configure: options => options.MaxRows = 3)))(document.AddSheet("Data", OfficeIMO.Excel.ExcelSheetNameValidationMode.Sanitize).Report()));
        Assert.Contains("3-row", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowsFromStopsNestedExpansionAtConfiguredLimit()
    {
        using var stream = new MemoryStream();
        using ExcelDocument document = ExcelDocument.Create(stream);
        var rows = new[]
        {
            new NestedRow
            {
                Name = "item",
                Values = new[]
                {
                    1,
                    2,
                    3
                }
            }
        };
        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
        {
            var sheet13 = document.AddSheet("Data");
            int nextRow14 = 1;
            var projection15 = new OfficeIMO.Data.ObjectFlattenerOptions();
            {
                (projection15).MaxRows = 2;
                (projection15).CollectionMode = CollectionMode.ExpandRows;
            }

            var table16 = sheet13.AddTable(rows, new ExcelTableOptions { Name = null, Style = ExcelTableStyle.TableStyleMedium2, IncludeAutoFilter = true, Projection = projection15 }, row: nextRow14);
            nextRow14 = table16.AsRange().LastRow + 1;
        });
        Assert.Contains("nested expansion", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowsFromHonorsPerCollectionLimitDuringNestedExpansion()
    {
        using var stream = new MemoryStream();
        using ExcelDocument document = ExcelDocument.Create(stream);
        var rows = new[]
        {
            new NestedEnumerableRow
            {
                Values = InfiniteValues()
            }
        };
        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
        {
            var sheet17 = document.AddSheet("Data");
            int nextRow18 = 1;
            var projection19 = new OfficeIMO.Data.ObjectFlattenerOptions();
            {
                (projection19).MaxRows = 100;
                (projection19).MaxCollectionItems = 2;
                (projection19).CollectionMode = CollectionMode.ExpandRows;
            }

            var table20 = sheet17.AddTable(rows, new ExcelTableOptions { Name = null, Style = ExcelTableStyle.TableStyleMedium2, IncludeAutoFilter = true, Projection = projection19 }, row: nextRow18);
            nextRow18 = table20.AsRange().LastRow + 1;
        });
        Assert.Contains("2-item", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowsFromSimpleFastPathRejectsColumnsBeyondConfiguredLimit()
    {
        using ExcelDocument document = ExcelDocument.Create();
        var sheet = document.AddSheet("Data");
        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => sheet.AddTable(new[] { new WideSimpleRow() }, new ExcelTableOptions { Projection = new ObjectFlattenerOptions { MaxColumns = 1 } }));
        Assert.Contains("1-column", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowsFromHonorsFinalCellLimitDuringNestedExpansion()
    {
        using var stream = new MemoryStream();
        using ExcelDocument document = ExcelDocument.Create(stream);
        var rows = new[]
        {
            new NestedRow
            {
                Name = "A",
                Values = new[]
                {
                    1,
                    2,
                    3
                }
            }
        };
        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
        {
            var sheet21 = document.AddSheet("Data");
            int nextRow22 = 1;
            var projection23 = new OfficeIMO.Data.ObjectFlattenerOptions();
            {
                (projection23).MaxCells = 5;
                (projection23).CollectionMode = CollectionMode.ExpandRows;
            }

            var table24 = sheet21.AddTable(rows, new ExcelTableOptions { Name = null, Style = ExcelTableStyle.TableStyleMedium2, IncludeAutoFilter = true, Projection = projection23 }, row: nextRow22);
            nextRow22 = table24.AsRange().LastRow + 1;
        });
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