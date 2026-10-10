using OfficeIMO.Pdf;
using OfficeIMO.Publisher;
using OfficeIMO.Reader.Publisher;
using OfficeIMO.TestSupport;
using System.Text.RegularExpressions;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderPublisherTableTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeCellsHaveGenericColumnsAndOneCanonicalTableAcrossRichAndChunkViews(bool transport) {
        OfficeDocumentReadResult result = Read(Fixture());
        if (transport) result = OfficeDocumentReadResultJson.Deserialize(result.ToJson());
        ReaderTable table = Assert.Single(result.EnumerateTables());
        Assert.Equal(new[] { "Column 1", "Column 2" }, table.Columns);
        Assert.Equal(new[] { "Table on page 2", "Top right", "P2 table left", "P2 table right", "Bottom Left", "Bottom Right" }, table.Rows.SelectMany(row => row));
        Assert.Equal(3, table.TotalRowCount); Assert.False(table.Truncated);
        Assert.Equal(2, table.Location!.Page); Assert.Equal(0, table.Location.TableIndex);
        Assert.Single(result.Tables); Assert.Single(result.Pages[1].Tables);
        Assert.Single(result.Chunks.SelectMany(chunk => chunk.Tables ?? Array.Empty<ReaderTable>()));
        OfficeDocumentBlock block = Assert.Single(result.Blocks, item => item.Kind == "table");
        Assert.Equal(table.Location.BlockAnchor, block.Id);
        Assert.Equal(2, block.Location.Page);
        Assert.Equal(block.Text, string.Concat(result.Chunks.Where(chunk => chunk.Location.LogicalOrder == block.Location.LogicalOrder).Select(chunk => chunk.Text)));
        string csv = table.ToCsv();
        Assert.StartsWith("Column 1,Column 2", csv);
        Assert.Contains("Table on page 2,Top right", csv);
    }

    [Theory]
    [InlineData(false, PdfProjectionPagePolicy.ContinuousFlow)]
    [InlineData(true, PdfProjectionPagePolicy.ContinuousFlow)]
    [InlineData(false, PdfProjectionPagePolicy.PreserveSourcePages)]
    [InlineData(true, PdfProjectionPagePolicy.PreserveSourcePages)]
    public void SemanticPdfDoesNotRepeatNativeTableTextAlongsideItsStoryBlock(bool transport, PdfProjectionPagePolicy policy) {
        OfficeDocumentReadResult result = Read(Fixture());
        if (transport) result = OfficeDocumentReadResultJson.Deserialize(result.ToJson());
        string text = PdfReadDocument.Open(result.ToPdfDocumentResult(new PdfProjectionOptions {
            PagePolicy = policy, IncludeMetadata = false
        }).ToBytes()).ExtractText();
        foreach (string cell in new[] { "Table on page 2", "Top right", "P2 table left", "P2 table right", "Bottom Left", "Bottom Right" })
            Assert.Single(Regex.Matches(text, Regex.Escape(cell)));
    }

    [Theory]
    [InlineData(0, 1)]
    [InlineData(-1, 1)]
    [InlineData(1, 1)]
    [InlineData(3, 3)]
    public void RowLimitAffectsStructuredRowsWithReportedLossAndRetainsCompleteStoryText(int maximum, int rows) {
        OfficeDocumentReadResult result = Read(Fixture(), new ReaderOptions { MaxTableRows = maximum });
        ReaderTable table = Assert.Single(result.Tables);
        Assert.Equal(rows, table.Rows.Count); Assert.Equal(3, table.TotalRowCount);
        Assert.Equal(rows < 3, table.Truncated);
        Assert.Contains("Bottom Right", result.Markdown);
        Assert.Contains(result.Blocks, block => block.Kind == "table" && block.Text!.Contains("Bottom Right"));
        if (!table.Truncated) return;
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "PUB_READER_TABLE_ROWS_TRUNCATED"
            && diagnostic.Attributes["lossKind"] == OfficeConversionLossKind.Omission.ToString());
        PdfDocumentConversionResult pdf = result.ToPdfDocumentResult(new PdfProjectionOptions { PagePolicy = PdfProjectionPagePolicy.ContinuousFlow });
        Assert.DoesNotContain("Bottom Right", PdfReadDocument.Open(pdf.ToBytes()).ExtractText());
        Assert.Contains(pdf.Warnings, warning => warning.Code == "PUB_READER_TABLE_ROWS_TRUNCATED" && warning.LossKind == OfficeConversionLossKind.Omission);
        Assert.Throws<InvalidOperationException>(() => pdf.RequireNoLoss());
    }

    [Fact]
    public void MergedCellsHaveOneAnchorValueAndAnExplicitFlatteningReport() {
        OfficeDocumentReadResult result = Read(PublisherTableFixture.MergeFirstRow(Fixture()));
        ReaderTable table = Assert.Single(result.Tables);
        Assert.Equal("Table on page 2\nTop right", table.Rows[0][0]); Assert.Empty(table.Rows[0][1]);
        Assert.Single(table.Rows.SelectMany(row => row).Where(value => value.Contains("Top right")));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "PUB_READER_TABLE_SPANS_FLATTENED"
            && diagnostic.Attributes["lossKind"] == OfficeConversionLossKind.Approximation.ToString());
        OfficeDocumentReadResult transported = OfficeDocumentReadResultJson.Deserialize(result.ToJson());
        Assert.Equal(table.Rows[0][0], Assert.Single(transported.EnumerateTables()).Rows[0][0]);
    }

    [Fact]
    public void UnknownCellTextIsNotProjectedAsAnEmptyDataset() {
        OfficeDocumentReadResult result = Read(PublisherTableFixture.WithoutTextMapping(Fixture()));
        Assert.Empty(result.Tables); Assert.All(result.Pages, page => Assert.Empty(page.Tables));
        Assert.Contains("Table on page 2", result.Markdown); Assert.Contains("Bottom Right", result.Markdown);
        Assert.DoesNotContain(result.Blocks, block => block.Kind == "table");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "PUB_TABLE_TEXT_MAPPING_UNRESOLVED");
    }

    [Fact]
    public void MasterTablesAreExtractedOnceWithoutInventingAPhysicalPageCitation() {
        byte[] original = Fixture(); PublisherDocument baseline = PublisherDocument.Load(original);
        PublisherPage master = Assert.Single(baseline.MasterPages);
        byte[] input = PublisherTableFixture.MoveToMaster(original, baseline.Pages[1].Id, master.Id);
        OfficeDocumentReadResult result = Read(input);
        ReaderTable table = Assert.Single(result.EnumerateTables());
        Assert.Null(table.Location!.Page); Assert.Equal("publisher-master-table", table.Location.SourceBlockKind);
        Assert.All(result.Pages, page => Assert.Empty(page.Tables));
        string text = PdfReadDocument.Open(result.ToPdfDocumentResult(new PdfProjectionOptions {
            PagePolicy = PdfProjectionPagePolicy.ContinuousFlow, IncludeMetadata = false
        }).ToBytes()).ExtractText();
        Assert.Single(Regex.Matches(text, "Table on page 2"));
    }

    private static OfficeDocumentReadResult Read(byte[] bytes, ReaderOptions? options = null) =>
        new OfficeDocumentReaderBuilder().AddPublisherHandler().Build().ReadDocument(bytes, "native.pub", options);
    private static byte[] Fixture() => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "PublisherFixtures", "Sample.pub"));
}
