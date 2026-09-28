using OfficeIMO.Reader.IWork;
using OfficeIMO.Reader.All;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderIWorkTests {
    [Theory]
    [InlineData("nim-iwork/simple.pages", "hello pages")]
    [InlineData("nim-iwork/simple.numbers", "a")]
    [InlineData("nim-iwork/simple.key", "hello keynote")]
    public void PublicReaderProjectsIndependentIWorkCorpusForPathsAndStreams(
        string relativePath, string marker) {
        string path = Fixture(relativePath);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build();

        OfficeDocumentReadResult fromPath = reader.ReadDocument(path);
        using FileStream input = File.OpenRead(path);
        OfficeDocumentReadResult fromStream = reader.ReadDocument(input, Path.GetFileName(path));

        Assert.Equal(ReaderInputKind.IWork, fromPath.Kind);
        Assert.Contains("officeimo.reader.iwork", fromPath.CapabilitiesUsed);
        Assert.Equal(fromPath.Chunks.Select(chunk => chunk.Text),
            fromStream.Chunks.Select(chunk => chunk.Text));
        Assert.Contains(marker, string.Join("\n", fromPath.Chunks.Select(chunk => chunk.Text)),
            StringComparison.OrdinalIgnoreCase);
        Assert.NotEmpty(fromPath.Pages);
        Assert.Equal(OfficeDocumentPageProvenance.LogicalContainer,
            fromPath.GetPageProvenance());
    }

    [Fact]
    public void NumbersTablesRetainCachedValuesAndReportBoundedProjection() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddIWorkHandler(new ReaderIWorkOptions { MaximumTableColumns = 2 })
            .Build();

        OfficeDocumentReadResult document = reader.ReadDocument(
            Fixture("numbers-parser/test-10-formulas.numbers"),
            new ReaderOptions { MaxTableRows = 2 });

        Assert.Equal(ReaderInputKind.IWork, document.Kind);
        Assert.True(document.Pages.Count >= 2);
        Assert.All(document.Tables, table => Assert.True(table.Rows.Count <= 2));
        Assert.Contains(document.Tables, table => table.Truncated);
        Assert.Contains(document.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_READER_TABLE_TRUNCATED");
        Assert.Contains(document.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_READER_FORMULA_CACHE");
    }

    [Fact]
    public void PagesTablesAndImagesRemainAvailableWhenBodyNeedsVisualFallback() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build();

        OfficeDocumentReadResult document = reader.ReadDocument(
            Fixture("picodocs/sample-v14.4.pages"));

        Assert.Equal(3, document.Tables.Count);
        Assert.Equal("Feature", document.Tables[0].Rows[0][0]);
        Assert.Contains(document.Tables[0].Rows, row => row.Contains("Preserve reading order"));
        OfficeDocumentAsset image = Assert.Single(document.Assets);
        Assert.Equal("image/png", image.MediaType);
        Assert.Null(image.PayloadBytes);
        Assert.Contains(document.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_TEXT_UNSUPPORTED");
    }

    [Fact]
    public void KeynoteProjectsSlidesNotesTablesAndImagePayloads() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddIWorkHandler(new ReaderIWorkOptions { IncludeImagePayloads = true })
            .Build();

        OfficeDocumentReadResult notes = reader.ReadDocument(Fixture("nim-iwork/simple.key"));
        OfficeDocumentReadResult table = reader.ReadDocument(
            Fixture("keynotekit/tabledeck-v15.2.1.key"));
        OfficeDocumentReadResult image = reader.ReadDocument(
            Fixture("keynotekit/imagedeck-v15.2.1.key"));

        Assert.Equal(2, notes.Pages.Count);
        Assert.Contains("note text here", notes.Markdown, StringComparison.Ordinal);
        Assert.Equal("Product", Assert.Single(table.Tables).Columns[0]);
        Assert.NotEmpty(Assert.Single(image.Assets).PayloadBytes!);
    }

    [Fact]
    public void AllPresetRoutesIWorkFilesThroughTheAdapter() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddAllOfficeIMOHandlers()
            .Build();
        OfficeDocumentReadResult result = reader.ReadDocument(Fixture("nim-iwork/simple.pages"));

        Assert.Equal(ReaderInputKind.IWork, result.Kind);
        Assert.Contains("officeimo.reader.iwork", result.CapabilitiesUsed);
    }

    [Theory]
    [InlineData("nim-iwork/simple.pages")]
    [InlineData("nim-iwork/simple.numbers")]
    [InlineData("nim-iwork/simple.key")]
    public void PreferContentRetainsValidatedIWorkRoutes(string relativePath) {
        string path = Fixture(relativePath);
        var options = new ReaderOptions { DetectionMode = ReaderDetectionMode.PreferContent };
        foreach (OfficeDocumentReader reader in new[] {
                     new OfficeDocumentReaderBuilder().AddIWorkHandler().Build(),
                     new OfficeDocumentReaderBuilder().AddAllOfficeIMOHandlers().Build()
                 }) {
            Assert.Equal(ReaderInputKind.IWork, reader.ReadDocument(path, options).Kind);
            using FileStream input = File.OpenRead(path);
            Assert.Equal(ReaderInputKind.IWork,
                reader.ReadDocument(input, Path.GetFileName(path), options).Kind);
        }
    }

    [Fact]
    public void ChunkOnlyReadCarriesSourceWarnings() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build();
        ReaderChunk[] chunks = reader.Read(Fixture("picodocs/sample-v14.4.pages")).ToArray();
        Assert.Contains(chunks.SelectMany(chunk => chunk.Warnings ?? Array.Empty<string>()),
            warning => warning.Contains("IWORK_PAGES_TEXT_UNSUPPORTED", StringComparison.Ordinal));
    }

    [Fact]
    public void TableRowBudgetCountsDataRowsAfterHeaders() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build();
        OfficeDocumentReadResult document = reader.ReadDocument(
            Fixture("keynotekit/tabledeck-v15.2.1.key"),
            new ReaderOptions { MaxTableRows = 1 });
        ReaderTable table = document.Tables[0];
        Assert.Single(table.Rows);
        Assert.Equal("Product", table.Columns[0]);
        Assert.Equal(2, table.TotalRowCount);
        Assert.True(table.Truncated);
    }

    [Fact]
    public void SplitTextKeepsOneLogicalMarkdownBlock() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build();
        OfficeDocumentReadResult document = reader.ReadDocument(
            Fixture("picodocs/sample-v14.4.pages"), new ReaderOptions { MaxChars = 256 });

        Assert.Contains(document.Chunks, chunk => chunk.ContinuesPreviousChunk);
        string markdown = Assert.IsType<string>(document.Markdown);
        Assert.Contains("Preserve reading order", markdown, StringComparison.Ordinal);
        Assert.Equal(1, markdown.Split("Preserve reading order", StringSplitOptions.None).Length - 1);
    }

    private static string Fixture(string relativePath) =>
        Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus",
            relativePath.Replace('/', Path.DirectorySeparatorChar));
}
