using OfficeIMO.Excel;
using OfficeIMO.Reader.Excel;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderExcelChunkingTests {
    [Fact]
    public void LongSheetMarkdownSplitsBetweenRowsWithoutDroppingAny() {
        const int dataRows = 300;
        string path = Path.Combine(Path.GetTempPath(), "reader-excel-chunking-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try {
            using (ExcelDocument document = ExcelDocument.Create()) {
                ExcelSheet sheet = document.AddSheet("Notes");
                sheet.Cell(1, 1).SetValue("Id");
                sheet.Cell(1, 2).SetValue("Note");
                for (int row = 2; row <= dataRows + 1; row++) {
                    sheet.Cell(row, 1).SetValue("row-" + (row - 1));
                    sheet.Cell(row, 2).SetValue("A note long enough to fill a chunk quickly | with a pipe " + new string('x', 40));
                }
                document.Save(path);
            }

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
}
