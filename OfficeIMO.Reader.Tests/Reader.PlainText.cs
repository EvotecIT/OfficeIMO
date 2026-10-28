using System;
using System.IO;
using System.Linq;
using System.Threading;
using OfficeIMO.Excel;
using OfficeIMO.Reader.Excel;
using OfficeIMO.Reader.Markdown;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderPlainTextTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MarkdownTablesRetainOneTextRepresentation(bool narrative) {
        string path = Path.Combine(Path.GetTempPath(), "reader-table-" + Guid.NewGuid().ToString("N") + ".md");
        string source = (narrative ? "Narrative before the table.\n\n" : "")
            + "| Item | Value |\n| --- | --- |\n| Delivery | 42 |";
        try {
            File.WriteAllText(path, source);
            var document = new OfficeDocumentReaderBuilder().AddMarkdownHandler().Build().ReadDocument(path);
            Assert.Single(document.EnumerateTables());
            string text = document.GetText();
            Assert.Equal(1, text.Split(new[] { "Delivery" }, StringSplitOptions.None).Length - 1);
            Assert.Contains("42", text);
            if (narrative) Assert.Contains("Narrative before the table.", text);
            else {
                Assert.Contains("Item\tValue", text);
                Assert.Contains("Delivery\t42", text);
            }
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(ReaderProjection.Full)]
    [InlineData(ReaderProjection.TextAndMetadata)]
    public void WorksheetTextDoesNotRepeatPromotedChunkTables(ReaderProjection projection) {
        using var workbook = ExcelDocument.Create();
        var sheet = workbook.AddWorksheet("Evidence");
        sheet.CellValue(1, 1, "Item"); sheet.CellValue(1, 2, "Value");
        sheet.CellValue(2, 1, "Delivery"); sheet.CellValue(2, 2, "42");
        var document = new OfficeDocumentReaderBuilder().AddExcelHandler().Build()
            .ReadDocument(workbook.ToBytes(), "evidence.xlsx", new ReaderOptions { Projection = projection });
        Assert.Single(document.EnumerateTables());
        Assert.Equal(1, document.GetText().Split(new[] { "Delivery" }, StringSplitOptions.None).Length - 1);
        Assert.Equal(document.GetText(), Assert.Single(document.GetPageText()).Text);
    }

    [Theory]
    [InlineData(8001, false)]
    [InlineData(16003, false)]
    [InlineData(8001, true)]
    [InlineData(16003, true)]
    public void TextRetainsSourceCharactersAcrossArtificialChunkBoundaries(int length, bool newline) {
        string path = Path.Combine(Path.GetTempPath(), "reader-text-" + Guid.NewGuid().ToString("N") + ".txt");
        string source = new string('a', 7999) + (newline ? "\r\n" : "ab") + new string('b', length - 8001) + "end";
        try {
            File.WriteAllText(path, source);
            var document = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().Build().ReadDocument(path);
            Assert.True(document.Chunks.Count > 1);
            // The plain-text reader normalizes CRLF to LF before bounded chunking.
            Assert.Equal(source.Replace("\r\n", "\n"), document.GetText());
        } finally { File.Delete(path); }
    }

    [Fact]
    public void TextKeepsSourceOrderAndRendersTheStructuredTableOnlyOnce() {
        var location = new ReaderLocation { Path = "evidence.pdf", Page = 1, BlockAnchor = "table-1", LogicalOrder = 1 };
        var table = new ReaderTable { Location = location, Columns = new[] { "Item", "Value" }, Rows = new[] { new[] { "Delivery", "42" } } };
        var heading = new OfficeDocumentBlock { Id = "heading", Text = "Evidence", Location = new ReaderLocation { Path = "evidence.pdf", Page = 1, LogicalOrder = 0 } };
        var tableBlock = new OfficeDocumentBlock { Id = "table-1", Kind = "table", Text = "Duplicate table prose", Location = location };
        var page = new OfficeDocumentPage { Number = 1, Blocks = new[] { heading, tableBlock }, Tables = new[] { table } };
        var document = new OfficeDocumentReadResult {
            Blocks = new[] { heading, tableBlock }, Tables = new[] { table }, Pages = new[] { page },
            Chunks = new[] { new ReaderChunk { Text = "Duplicate chunk projection", Tables = new[] { table } } }
        };
        Assert.Equal("Evidence" + Environment.NewLine + "Item\tValue" + Environment.NewLine + "Delivery\t42", document.GetText());
        var projected = Assert.Single(document.GetPageText());
        Assert.Same(page, projected.Page);
        Assert.Equal(document.GetText(), projected.Text);
        Assert.DoesNotContain("Duplicate", projected.Text);
    }

    [Fact]
    public void ChunkOnlyResultsKeepTheirTextWithoutDuplicatingTheirTables() {
        var table = new ReaderTable { Columns = new[] { "Item" }, Rows = new[] { new[] { "Delivery" } } };
        var document = new OfficeDocumentReadResult {
            Chunks = new[] { new ReaderChunk { Text = "A narrative with the extracted Delivery table", Tables = new[] { table } } },
            Tables = new[] { table }
        };
        Assert.Equal("A narrative with the extracted Delivery table", document.GetText());
    }

    [Fact]
    public void PageTextKeepsPageBoundariesAndDoesNotInventUnobservedPages() {
        var document = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Pdf,
            Pages = new[] {
                new OfficeDocumentPage { Number = 2, Blocks = new[] { new OfficeDocumentBlock { Text = "Second page" } } },
                new OfficeDocumentPage { Number = 5, Blocks = new[] { new OfficeDocumentBlock { Text = "Fifth page" } } }
            }
        };
        var pages = document.GetPageText();
        Assert.Equal(new[] { "Second page", "Fifth page" }, pages.Select(page => page.Text));
        Assert.Equal(new int?[] { 2, 5 }, pages.Select(page => page.Page.Number));
        Assert.All(pages, page => Assert.Equal(document.GetPageProvenance(), page.Provenance));
        Assert.Empty(new OfficeDocumentReadResult { Blocks = new[] { new OfficeDocumentBlock { Text = "Unpaginated" } } }.GetPageText());
    }

    [Fact]
    public void TextProjectionHonoursCancellation() {
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        var document = new OfficeDocumentReadResult { Blocks = new[] { new OfficeDocumentBlock { Text = "Evidence" } } };
        Assert.Throws<OperationCanceledException>(() => document.GetText(cancelled.Token));
        Assert.Throws<OperationCanceledException>(() => new OfficeDocumentReadResult().GetText(cancelled.Token));
        Assert.Throws<OperationCanceledException>(() => new OfficeDocumentReadResult().GetPageText(cancelled.Token));
    }
}
