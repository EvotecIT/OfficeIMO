using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderPdfPlainTextTests {
    [Theory]
    [InlineData(ReaderProjection.Full, true, 8000)]
    [InlineData(ReaderProjection.Full, false, 13)]
    [InlineData(ReaderProjection.TextAndMetadata, true, 8000)]
    [InlineData(ReaderProjection.TextAndMetadata, false, 13)]
    public void TableTextRetainsLiteralCellsAcrossProjectionAndChunkBoundaries(
        ReaderProjection projection, bool chunkByPage, int maxChars) {
        using var pdf = PdfDocument.Create(builder => builder.Content(content => content.AddTable(new[] {
            new[] { "Name", "Tags" }, new[] { "Alpha", "One | Two" }
        }, style: new PdfTableStyle { HeaderRowCount = 1 })));
        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(pdf.ToBytes(), "literal-table.pdf",
            new ReaderOptions { Projection = projection, MaxChars = maxChars },
            new ReaderPdfOptions { ChunkByPage = chunkByPage });

        string text = result.GetText();
        Assert.Equal("Name\tTags" + Environment.NewLine + "Alpha\tOne | Two", text);
        Assert.All(result.Chunks, chunk => Assert.True((chunk.Text?.Length ?? 0) <= maxChars));
        if (projection == ReaderProjection.Full) {
            Assert.Equal(text, Assert.Single(result.GetPageText()).Text);
            Assert.Contains("One \\| Two", result.Markdown);
            Assert.Contains("\\|", string.Join("", result.Chunks.Select(chunk => chunk.Markdown)));
            Assert.Empty(Assert.Single(result.Blocks, block => block.Kind == "table").Text);
        } else {
            Assert.Null(result.Markdown);
            Assert.Empty(result.Tables);
        }
    }
}
