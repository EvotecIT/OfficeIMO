using OfficeIMO.Reader.IWork;
using OfficeIMO.Reader.All;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderIWorkTests {
    [Theory]
    [InlineData("sample.pages", "application/vnd.apple.pages")]
    [InlineData("sample.numbers", "application/vnd.apple.numbers")]
    [InlineData("sample.key", "application/vnd.apple.keynote")]
    public void IWorkExtensionDetectionReportsItsRegisteredMediaType(
        string sourceName, string expectedMediaType) {
        ReaderDetectionResult detection = new OfficeDocumentReaderBuilder()
            .AddIWorkHandler().Build().Detect(Array.Empty<byte>(), sourceName,
                new ReaderDetectionOptions { Mode = ReaderDetectionMode.ExtensionOnly });

        Assert.Equal(ReaderInputKind.IWork, detection.Kind);
        Assert.Equal(expectedMediaType, detection.MediaType);
    }

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
        OfficeDocumentBlock tableBlock = Assert.Single(document.Blocks,
            block => block.Kind == "table" && block.Text.Contains("Feature", StringComparison.Ordinal));
        Assert.Equal(DocumentReaderEngine.BuildRichTableText(document.Tables[0]), tableBlock.Text);
        Assert.NotEqual(document.Tables[0].ToMarkdownTable(), tableBlock.Text);
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
        Assert.Contains(notes.Blocks, block => block.Location.SourceBlockKind == "presenter-notes"
            && block.Text.Contains("note text here", StringComparison.Ordinal));
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
    public async Task PreferContentRetainsValidatedIWorkRoutes(string relativePath) {
        string path = Fixture(relativePath);
        var options = new ReaderOptions { DetectionMode = ReaderDetectionMode.PreferContent };
        foreach (OfficeDocumentReader reader in new[] {
                     new OfficeDocumentReaderBuilder().AddIWorkHandler().Build(),
                     new OfficeDocumentReaderBuilder().AddAllOfficeIMOHandlers().Build()
                 }) {
            OfficeDocumentReadResult fromPath = reader.ReadDocument(path, options);
            Assert.Equal(ReaderInputKind.IWork, fromPath.Kind);
            Assert.DoesNotContain(fromPath.Diagnostics,
                diagnostic => diagnostic.Code == "input-kind-mismatch");
            using FileStream input = File.OpenRead(path);
            OfficeDocumentReadResult fromStream = reader.ReadDocument(input, Path.GetFileName(path), options);
            Assert.Equal(ReaderInputKind.IWork, fromStream.Kind);
            Assert.DoesNotContain(fromStream.Diagnostics,
                diagnostic => diagnostic.Code == "input-kind-mismatch");
            OfficeDocumentReadResult asyncPath = await reader.ReadDocumentAsync(path, options);
            Assert.Equal(ReaderInputKind.IWork, asyncPath.Kind);
            Assert.DoesNotContain(asyncPath.Diagnostics,
                diagnostic => diagnostic.Code == "input-kind-mismatch");
            using FileStream asyncInput = File.OpenRead(path);
            OfficeDocumentReadResult asyncStream = await reader.ReadDocumentAsync(
                asyncInput, Path.GetFileName(path), options);
            Assert.Equal(ReaderInputKind.IWork, asyncStream.Kind);
            Assert.DoesNotContain(asyncStream.Diagnostics,
                diagnostic => diagnostic.Code == "input-kind-mismatch");
        }
    }

    [Fact]
    public void PreferContentUsesTheKnownIWorkPackageLimitBeforeDetection() {
        const long packageLimit = 512L * 1024L * 1024L;
        var registry = new ReaderHandlerRegistry();
        registry.Register(new ReaderHandlerRegistration {
            Id = "officeimo.tests.iwork-limit",
            Kind = ReaderInputKind.IWork,
            Extensions = new[] { ".pages" },
            DefaultMaxInputBytes = packageLimit,
            MaxInputBytesCeiling = packageLimit,
            ReadPath = (_, _, _) => Array.Empty<ReaderChunk>(),
            ReadStream = (_, _, _, _) => Array.Empty<ReaderChunk>()
        }, replaceExisting: false);
        using (DocumentReaderEngine.UseHandlerRegistry(registry.CaptureSnapshot())) {
            var options = new ReaderOptions { DetectionMode = ReaderDetectionMode.PreferContent };
            Assert.Equal(packageLimit, DocumentReaderEngine.ResolveInitialMaxInputBytes(
                "large.pages", options));
            Assert.Equal(packageLimit, DocumentReaderEngine.ResolveStreamMaxInputBytes(
                "large.pages", options, streamCanSeek: false));
            Assert.Equal(packageLimit, DocumentReaderEngine.ResolveStreamMaxInputBytes(
                "large.pages", options, streamCanSeek: true));
            Assert.Equal(64L * 1024L * 1024L, DocumentReaderEngine.ResolveStreamMaxInputBytes(
                "unknown.bin", options, streamCanSeek: true));
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
        Assert.Equal(1, markdown.Split(new[] { "Preserve reading order" }, StringSplitOptions.None).Length - 1);
    }

    [Fact]
    public void SplitTableMarkdownRespectsChunkBudgetAndReassembles() {
        const int maxChars = 256;
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build();
        OfficeDocumentReadResult document = reader.ReadDocument(
            Fixture("picodocs/sample-v14.4.pages"), new ReaderOptions { MaxChars = maxChars });

        ReaderChunk first = Assert.Single(document.Chunks,
            chunk => chunk.Tables?.Count > 0 && chunk.Tables[0] == document.Tables[0]);
        string? anchor = first.Location.BlockAnchor;
        ReaderChunk[] parts = document.Chunks.Where(chunk => chunk.Location.BlockAnchor == anchor).ToArray();
        Assert.True(parts.Length > 1);
        Assert.All(parts, part => Assert.InRange(part.Markdown!.Length, 0, maxChars));
        Assert.Equal(document.Tables[0].ToMarkdownTable(), string.Concat(parts.Select(part => part.Markdown)));
    }

    [Fact]
    public void IWorkResultsKeepTheirNeutralFormatIdentity() {
        Assert.Equal(OfficeDocumentFormat.IWork,
            OfficeDocumentReadResultPdfExtensions.MapFormat(ReaderInputKind.IWork));
    }

    private static string Fixture(string relativePath) =>
        Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus",
            relativePath.Replace('/', Path.DirectorySeparatorChar));
}
