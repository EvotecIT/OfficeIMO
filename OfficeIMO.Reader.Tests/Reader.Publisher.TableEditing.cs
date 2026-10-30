using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Publisher;
using OfficeIMO.Reader.Publisher;
using OfficeIMO.Tests;
using System.Text.RegularExpressions;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderPublisherTableEditingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EditedNativeCellsReachStructuredTransportAndSemanticPdfWithoutStaleStoryText(bool transport) {
        string path = NativePath("TableLabel", "native-tablelabel.pub");
        byte[] source = File.ReadAllBytes(path);
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable original = Assert.Single(document.Pages[0].Tables);
        PublisherTextStory originalStory = Assert.Single(document.TextStories, story => story.Id == original.StoryId);
        string oldStoryText = originalStory.Text;
        var sourceReport = document.SourceReadReport;
        document.ReplaceTableCellParagraphs(original.Id, 0, 0, new[] {
            Paragraph("Expanded inventory", bold: true), Paragraph("Second paragraph")
        });
        document.SetTableBounds(original.Id, 18, 66, 198, 90);
        PublisherTable current = Assert.Single(document.Pages[0].Tables);
        Assert.Equal("Item", original.Cells[0].Text);
        Assert.Equal("Expanded inventory\nSecond paragraph", current.Cells[0].Text);
        Assert.Equal(original.Cells.Skip(1).Select(cell => cell.Text), current.Cells.Skip(1).Select(cell => cell.Text));
        Assert.Equal(oldStoryText, originalStory.Text);
        Assert.Same(sourceReport, document.SourceReadReport);
        OfficeDocumentReadResult result = document.ToOfficeDocumentReadResult("edited-table.pub", new ReaderOptions { ComputeHashes = true });
        if (transport) result = OfficeDocumentReadResultJson.Deserialize(result.ToJson());
        ReaderTable table = Assert.Single(result.EnumerateTables());
        Assert.Equal(new[] { "Expanded inventory\nSecond paragraph", "Count", "Alpha", "2" }, table.Rows.SelectMany(row => row));
        string semanticText = string.Concat(result.Blocks.Select(block => block.Text));
        Assert.Single(Regex.Matches(semanticText, "Expanded inventory"));
        Assert.Single(Regex.Matches(semanticText, "Second paragraph"));
        var pdfOptions = new PdfOptions();
        pdfOptions.UseTextFallbacks();
        string pdfText = PdfReadDocument.Open(result.ToPdfDocumentResult(new PdfProjectionOptions {
            PagePolicy = PdfProjectionPagePolicy.ContinuousFlow, IncludeMetadata = false, PdfOptions = pdfOptions
        }).ToBytes()).ExtractText();
        Assert.Single(Regex.Matches(pdfText, "Expanded inventory"));
        Assert.Single(Regex.Matches(pdfText, "Second paragraph"));
        Assert.Null(result.Source.SourceHash); Assert.Null(result.Source.LengthBytes);
        Assert.Equal(OfficeDocumentPageProvenance.Native, result.GetPageProvenance());

        document.ReplaceTableCellParagraphs(original.Id, 0, 0, Array.Empty<OfficeRichTextParagraph>());
        document.ReplaceTableCellParagraphs(original.Id, 0, 0, new[] { Paragraph("Refilled") });
        ReaderTable refilled = Assert.Single(document.ToOfficeDocumentReadResult("edited-table.pub").EnumerateTables());
        Assert.Equal("Refilled", refilled.Rows[0][0]);
        Assert.Equal("Expanded inventory\nSecond paragraph", table.Rows[0][0]);
        Assert.Equal(source, File.ReadAllBytes(path));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AuthoredTableHasOneStoryAndOneCurrentTableAcrossReaderViews(bool transport) {
        string path = NativePath("TextBlank", "native-textblank.pub");
        byte[] source = File.ReadAllBytes(path);
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTextStory originalStory = Assert.Single(document.TextStories);
        var sourceReport = document.SourceReadReport;
        PublisherTable created = document.AddTable(document.Pages[1].Id, 72, 72, new[] { 120D, 90D }, new[] { 30D, 30D });
        document.ReplaceTableCellParagraphs(created.Id, 0, 0, new[] { Paragraph("Authored item") });
        document.ReplaceTableCellParagraphs(created.Id, 0, 1, new[] { Paragraph("Quantity") });
        document.ReplaceTableCellParagraphs(created.Id, 1, 0, new[] { Paragraph("Widget") });
        document.ReplaceTableCellParagraphs(created.Id, 1, 1, new[] { Paragraph("7") });
        Assert.Equal(2, document.TextStories.Count);
        Assert.NotEqual(originalStory.Id, created.StoryId);
        Assert.Equal("Publisher text\n", originalStory.Text);
        OfficeDocumentReadResult result = document.ToOfficeDocumentReadResult("authored-table.pub", new ReaderOptions { ComputeHashes = true, MaxChars = 256 });
        if (transport) result = OfficeDocumentReadResultJson.Deserialize(result.ToJson());
        ReaderTable table = Assert.Single(result.EnumerateTables());
        ReaderTable chunkTable = Assert.Single(DocumentReaderEngine.ExtractTables(result.Chunks));
        Assert.Equal(new[] { "Authored item", "Quantity", "Widget", "7" }, table.Rows.SelectMany(row => row));
        Assert.Equal(table.Rows.SelectMany(row => row), chunkTable.Rows.SelectMany(row => row));
        Assert.Equal(2, table.Location!.Page);
        Assert.Single(DocumentReaderEngine.ExportTables(DocumentReaderEngine.ExtractTables(result.Chunks)));
        string semanticText = string.Concat(result.Blocks.Select(block => block.Text));
        Assert.Single(Regex.Matches(semanticText, "Authored item"));
        Assert.Single(Regex.Matches(semanticText, "Publisher text"));
        Assert.Null(result.Source.SourceHash); Assert.Null(result.Source.LengthBytes);
        Assert.Same(sourceReport, document.SourceReadReport);
        Assert.Equal(source, File.ReadAllBytes(path));
    }

    private static OfficeRichTextParagraph Paragraph(string text, bool bold = false) =>
        new(new[] { new OfficeRichTextRun(text, 10, OfficeColor.FromRgb(0, 90, 170), bold: bold, fontFamily: "Arial") });

    private static string NativePath(string fixture, string name) => Path.Combine(RepositoryTestPaths.Find(),
        "OfficeIMO.Publisher.Tests", "Fixtures", "NativeComparison", fixture, name);
}
