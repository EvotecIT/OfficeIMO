using System;
using System.IO;
using System.Linq;
using OfficeIMO.Excel;
using OfficeIMO.Reader.Excel;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderExcelSheetSelectionTests {
    [Theory]
    [InlineData("path", ReaderProjection.Full)]
    [InlineData("stream", ReaderProjection.Full)]
    [InlineData("loaded", ReaderProjection.Full)]
    [InlineData("incremental", ReaderProjection.Full)]
    [InlineData("path", ReaderProjection.TextAndMetadata)]
    [InlineData("stream", ReaderProjection.TextAndMetadata)]
    [InlineData("loaded", ReaderProjection.TextAndMetadata)]
    [InlineData("incremental", ReaderProjection.TextAndMetadata)]
    public void OneBasedSelectionAppliesAcrossEveryProjectionWithoutMutatingThePolicy(string route, ReaderProjection projection) {
        var format = new ReaderExcelOptions { SheetNumber = 2, HeadersInFirstRow = false };
        var reader = new OfficeDocumentReaderBuilder().AddExcelHandler(format).Build();
        string path = Path.Combine(Path.GetTempPath(), "reader-sheet-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try {
            using (var document = ExcelDocument.Create()) {
                document.AddSheet("First").CellValue(1, 1, "Excluded first value");
                document.AddSheet("2").CellValue(1, 1, "Selected second value");
                document.Save(path);
            }
            var options = new ReaderOptions { Projection = projection };
            if (route == "incremental") {
                var chunks = reader.EnumerateChunks(path, options).ToArray();
                Assert.NotEmpty(chunks);
                Assert.All(chunks, chunk => Assert.Equal("2", chunk.Location!.Sheet));
                Assert.Contains(chunks, chunk => chunk.Text!.Contains("Selected second value"));
            } else {
                OfficeDocumentReadResult result;
                if (route == "loaded") {
                    using var document = ExcelDocument.Load(path);
                    result = document.ReadDocument(path, options, format);
                    Assert.Equal(2, document.Sheets.Count);
                } else if (route == "stream") {
                    using var stream = File.OpenRead(path);
                    result = reader.ReadDocument(stream, "selection.xlsx", options);
                } else result = reader.ReadDocument(path, options);
                Assert.NotEmpty(result.Chunks);
                Assert.All(result.Chunks, chunk => Assert.Equal("2", chunk.Location!.Sheet));
                Assert.Contains(result.Chunks, chunk => chunk.Text!.Contains("Selected second value"));
                Assert.DoesNotContain(result.EnumerateBlocks(), block => block.Location?.Sheet == "First");
                Assert.All(result.Pages, page => Assert.Equal("2", page.Name));
            }
            Assert.Equal(2, format.SheetNumber);
            Assert.Null(format.SheetName);
        } finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void RegistrationRejectsContradictoryAndNonPositiveSelections() {
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeDocumentReaderBuilder().AddExcelHandler(new ReaderExcelOptions { SheetNumber = 0 }));
        Assert.Throws<ArgumentException>(() => new OfficeDocumentReaderBuilder().AddExcelHandler(new ReaderExcelOptions { SheetNumber = 1, SheetName = "Evidence" }));
    }

    [Fact]
    public void RegistrationSnapshotsSelectionAndResolvesItForEachWorkbook() {
        var format = new ReaderExcelOptions { SheetNumber = 2, HeadersInFirstRow = false };
        var reader = new OfficeDocumentReaderBuilder().AddExcelHandler(format).Build();
        format.SheetNumber = 1;
        foreach (string selectedName in new[] { "2", "Evidence" }) {
            using var document = ExcelDocument.Create();
            document.AddSheet("First").CellValue(1, 1, "Excluded");
            document.AddSheet(selectedName).CellValue(1, 1, "Selected");
            var result = reader.ReadDocument(document.ToBytes(), "source.xlsx");
            Assert.NotEmpty(result.Chunks);
            Assert.All(result.Chunks, chunk => Assert.Equal(selectedName, chunk.Location!.Sheet));
        }
    }

    [Fact]
    public void APositionOutsideTheWorkbookFailsInsteadOfReadingAllSheets() {
        using var document = ExcelDocument.Create();
        document.AddSheet("Evidence").CellValue(1, 1, "Keep private");
        var format = new ReaderExcelOptions { SheetNumber = 2 };
        Assert.Throws<ArgumentOutOfRangeException>(() => document.ReadDocument("source.xlsx", formatOptions: format));
        Assert.Equal(2, format.SheetNumber);
    }
}
