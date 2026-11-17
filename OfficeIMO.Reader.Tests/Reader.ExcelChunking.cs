using OfficeIMO.Excel;
using OfficeIMO.Reader.Excel;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderExcelChunkingTests {
    private const int DataRows = 300;

    [Fact]
    public void LongSheetMarkdownSplitsBetweenRowsWithoutDroppingAny() {
        const int dataRows = DataRows;
        string path = CreateNotesWorkbook();
        try {
            OfficeDocumentReadResult result = new OfficeDocumentReaderBuilder().AddExcelHandler().Build()
                .ReadDocument(path, new ReaderOptions { MaxChars = 2_000 });

            ReaderChunk[] chunks = result.Chunks.Where(chunk => chunk.Tables is { Count: > 0 }).ToArray();
            Assert.True(chunks.Length > 1);
            Assert.All(chunks, chunk => {
                Assert.True(chunk.Markdown!.Length <= 2_000, chunk.Markdown.Length.ToString());
                Assert.DoesNotContain("truncated", chunk.Markdown, StringComparison.Ordinal);
                Assert.StartsWith("### Notes", chunk.Markdown, StringComparison.Ordinal);
            });
            string[] ids = chunks.SelectMany(chunk => chunk.Tables![0].Rows).Select(row => row[0]).ToArray();
            Assert.Equal(Enumerable.Range(1, dataRows).Select(index => "row-" + index), ids);
            Assert.All(Enumerable.Range(1, dataRows), index =>
                Assert.Contains(chunks, chunk => chunk.Markdown!.Contains("| row-" + index + " |", StringComparison.Ordinal)));
            Assert.Equal(chunks.Length, chunks.Select(chunk => chunk.Id).Distinct().Count());
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void SplittingABlockCappedByMaxTableRowsKeepsTheOmittedRowCount() {
        string path = CreateNotesWorkbook();
        try {
            OfficeDocumentReadResult result = new OfficeDocumentReaderBuilder().AddExcelHandler().Build()
                .ReadDocument(path, new ReaderOptions { MaxChars = 2_000, MaxTableRows = 50 });

            ReaderTable[] tables = result.Chunks.Where(chunk => chunk.Tables is { Count: > 0 }).Select(chunk => chunk.Tables![0]).ToArray();
            Assert.True(tables.Length > 2);
            Assert.Equal(100, tables.Sum(table => table.Rows.Count));
            Assert.Equal(DataRows, tables.Sum(table => table.TotalRowCount));
            Assert.Equal(2, tables.Count(table => table.Truncated));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void TruncationNoteCountsTowardMaxChars() {
        // A single retained row fits MaxChars on its own, but not together with the truncation note.
        string path = Path.Combine(Path.GetTempPath(), "reader-excel-note-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try {
            using (ExcelDocument document = ExcelDocument.Create()) {
                ExcelSheet sheet = document.AddSheet("Notes");
                sheet.Cell(1, 1).SetValue("Id");
                for (int row = 2; row <= 4; row++) sheet.Cell(row, 1).SetValue(new string((char)('a' + row), 210));
                document.Save(path);
            }

            OfficeDocumentReadResult result = new OfficeDocumentReaderBuilder().AddExcelHandler().Build()
                .ReadDocument(path, new ReaderOptions { MaxChars = 256, MaxTableRows = 1 });

            ReaderChunk[] chunks = result.Chunks.Where(chunk => chunk.Tables is { Count: > 0 }).ToArray();
            Assert.Contains(chunks, chunk => chunk.Markdown!.Contains("<!-- truncated -->", StringComparison.Ordinal));
            Assert.All(chunks, chunk => Assert.True(chunk.Markdown!.Length <= 256
                || chunk.Warnings!.Any(warning => warning.Contains("exceeds MaxChars", StringComparison.Ordinal)),
                chunk.Markdown.Length.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private static string CreateNotesWorkbook() {
        string path = Path.Combine(Path.GetTempPath(), "reader-excel-chunking-" + Guid.NewGuid().ToString("N") + ".xlsx");
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Notes");
        sheet.Cell(1, 1).SetValue("Id");
        sheet.Cell(1, 2).SetValue("Note");
        for (int row = 2; row <= DataRows + 1; row++) {
            sheet.Cell(row, 1).SetValue("row-" + (row - 1));
            sheet.Cell(row, 2).SetValue("A note long enough to fill a chunk quickly | with a pipe " + new string('x', 40));
        }
        document.Save(path);
        return path;
    }
}
