using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed class PdfEvidenceProjectionTests {
    [Theory]
    [InlineData(true, 8000, false)]
    [InlineData(false, 8000, false)]
    [InlineData(true, 32, false)]
    [InlineData(false, 32, false)]
    [InlineData(true, 32, true)]
    [InlineData(false, 32, true)]
    public void TableOnlyPdfCapturesEachCellOnceWithinItsEvidenceBudget(bool perPage, int chunkSize, bool roundTrip) {
        byte[] source = TablePdf();
        var reader = new OfficeDocumentReaderBuilder().AddPdfHandler(new() { ChunkByPage = perPage }).Build();
        var result = reader.ReadDocument(source, "table.pdf", new() { MaxChars = chunkSize });
        if (roundTrip) result = OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(result))!;
        Assert.NotEmpty(result.Chunks);
        Assert.NotEmpty(result.Tables);
        Assert.All(result.EnumerateBlocks(), block => Assert.True(string.IsNullOrWhiteSpace(block.Text)));
        var expectedRows = result.EnumerateTables().Sum(table => table.Rows.Count);
        var document = OfficeAiDocument.FromReadResult(source, result);
        Assert.Equal(expectedRows, document.Evidence.Count(item => item.Kind == "table-row"));
        Assert.Single(document.Evidence, item => item.Text.Contains("A-100", StringComparison.Ordinal));
        Assert.Single(document.Evidence, item => item.Text.Contains("B-200", StringComparison.Ordinal));
        int structuredCharacters = document.Evidence.Where(item => item.Kind is "table" or "table-row").Sum(item => item.Text.Length);
        var bounded = OfficeAiDocument.FromReadResult(source, result, limits: new() { MaxDocumentCharacters = structuredCharacters });
        Assert.Equal(structuredCharacters, bounded.Evidence.Sum(item => item.Text.Length));
    }

    [Fact]
    public void PdfChunkTextWithoutACompleteBlockProjectionIsRetained() {
        var result = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Pdf,
            Chunks = [new() { Kind = ReaderInputKind.Pdf, Text = "Source note outside the table.", Location = new() { Page = 1, SourceBlockKind = "page" },
                Tables = [new() { Columns = ["Code"], Rows = [["A-100"]] }] }]
        };
        var document = OfficeAiDocument.FromReadResult([1], result);
        Assert.Contains(document.Evidence, item => item.Text == "Source note outside the table.");
        Assert.Contains(document.Evidence, item => item.Kind == "table-row" && item.Text.Contains("A-100"));
    }

    [Theory]
    [InlineData(1, "page", false)]
    [InlineData(2, "page", false)]
    [InlineData(null, "document", false)]
    [InlineData(1, "page", true)]
    [InlineData(2, "page", true)]
    [InlineData(null, "document", true)]
    public void PartialPdfProjectionKeepsNarrativeOutsideItsStructuredTable(int? page, string sourceKind, bool roundTrip) {
        var result = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Pdf,
            Blocks = [new() { Kind = "table", Text = "", Location = new() { Page = 1 } }],
            Tables = [new() { Columns = ["Code"], Rows = [["A-100"]], Location = new() { Page = 1 } }],
            Chunks = [new() { Kind = ReaderInputKind.Pdf, Text = "Independent narrative source.",
                Location = new() { Page = page, SourceBlockKind = sourceKind } }]
        };
        if (roundTrip) result = OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(result))!;
        var document = OfficeAiDocument.FromReadResult([1], result);
        Assert.Contains(document.Evidence, item => item.Text == "Independent narrative source.");
        Assert.Contains(document.Evidence, item => item.Kind == "table-row" && item.Text.Contains("A-100"));
    }

    [Fact]
    public void TableChunkFallbackRemainsWhenStructuredScopeIsIncomplete() {
        var result = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Pdf,
            Tables = [new() { Columns = ["Code"], Rows = [["A-100"]], Location = new() { Path = "table.pdf", Page = 1, SourceBlockIndex = 0 } }],
            Chunks = [new() { Kind = ReaderInputKind.Pdf, Text = "Unprojected second table B-200.",
                Location = new() { Path = "table.pdf", Page = 1, SourceBlockIndex = 0, SourceBlockKind = "table" },
                Diagnostics = new() { TableCount = 2 } }]
        };
        var document = OfficeAiDocument.FromReadResult([1], result);
        Assert.Contains(document.Evidence, item => item.Text.Contains("B-200"));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void TableTruncationCannotBeBypassedThroughMarkdownFallback(bool perPage) {
        byte[] source = TablePdf();
        var reader = new OfficeDocumentReaderBuilder().AddPdfHandler(new() { ChunkByPage = perPage }).Build();
        var result = reader.ReadDocument(source, "table.pdf", new() { MaxTableRows = 1, MaxChars = 32 });
        var document = OfficeAiDocument.FromReadResult(source, result);
        Assert.True(document.HasSourceDiagnostics);
        Assert.Single(document.Evidence, item => item.Kind == "table-row");
        Assert.Contains(document.Evidence, item => item.Text.Contains("A-100"));
        Assert.DoesNotContain(document.Evidence, item => item.Text.Contains("B-200"));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void TextAroundTablesRemainsSeparateCitableEvidence(bool perPage) {
        byte[] source = PdfDocument.Create(builder => builder.Content(content => {
            content.Text("Inventory introduction.");
            content.Table(new[] { new[] { "Code", "Count" }, new[] { "A-100", "42" }, new[] { "B-200", "7" } },
                style: new PdfTableStyle { HeaderRowCount = 1 });
            content.Text("Inventory conclusion.");
        })).ToBytes();
        var reader = new OfficeDocumentReaderBuilder().AddPdfHandler(new() { ChunkByPage = perPage }).Build();
        var document = OfficeAiDocument.FromReadResult(source, reader.ReadDocument(source, "inventory.pdf"));
        Assert.Single(document.Evidence, item => item.Text.Contains("Inventory introduction."));
        Assert.Single(document.Evidence, item => item.Text.Contains("Inventory conclusion."));
        Assert.Single(document.Evidence, item => item.Text.Contains("A-100"));
        Assert.Single(document.Evidence, item => item.Text.Contains("B-200"));
    }

    private static byte[] TablePdf() => PdfDocument.Create(builder => builder.Content(content => content.Table(new[] {
        new[] { "Code", "Name", "Qty" }, new[] { "A-100", "Alpha", "2" }, new[] { "B-200", "Beta", "14" }
    }, style: new PdfTableStyle { HeaderRowCount = 1 }))).ToBytes();
}
