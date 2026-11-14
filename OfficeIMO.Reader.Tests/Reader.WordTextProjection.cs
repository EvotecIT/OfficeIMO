using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Word;
using OfficeIMO.Word;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public sealed class ReaderWordTextProjectionTests {
    private const string StreamedCapability = "officeimo.reader.word.streamed-text";

    // A running header stops the streamed pass, so that row reads the Word model and reduces it to text.
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TextProjectionReturnsOnlyTextChunksWithTheTextOfTheLeanRead(bool withHeader) {
        byte[] package = CreateDocument(withHeader);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddWordHandler().Build();

        OfficeDocumentReadResult lean = reader.ReadDocument(package, "report.docx",
            new ReaderOptions { Projection = ReaderProjection.TextAndMetadata });
        OfficeDocumentReadResult text = reader.ReadDocument(package, "report.docx",
            new ReaderOptions { Projection = ReaderProjection.Text });

        Assert.NotEmpty(lean.Tables);
        Assert.Equal(lean.GetText(), text.GetText());
        Assert.NotEmpty(text.Chunks);
        Assert.All(text.Chunks, chunk => {
            Assert.False(string.IsNullOrEmpty(chunk.Text));
            Assert.Null(chunk.Tables);
            Assert.Null(chunk.Markdown);
        });
        Assert.Empty(text.Blocks);
        Assert.Empty(text.Tables);
        Assert.Empty(text.Pages);
        Assert.Empty(text.Links);
        Assert.Empty(text.Metadata);
        Assert.Null(text.Markdown);
        Assert.Contains("officeimo.reader.projection.text", text.CapabilitiesUsed);
        Assert.Equal(!withHeader, text.CapabilitiesUsed.Contains(StreamedCapability));
        Assert.DoesNotContain(StreamedCapability, lean.CapabilitiesUsed);
    }

    // Table shapes whose streamed text must still match the model: a grid wider than every row, a grid with no
    // rows, and a header row whose revised properties record it as an ordinary row.
    [Theory]
    [InlineData("<w:tblGrid><w:gridCol/><w:gridCol/><w:gridCol/></w:tblGrid>"
        + "<w:tr><w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>1,200</w:t></w:r></w:p></w:tc></w:tr>", true)]
    [InlineData("<w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>", true)]
    [InlineData("<w:tblGrid><w:gridCol/></w:tblGrid>"
        + "<w:tr><w:trPr><w:tblHeader/><w:trPrChange w:id=\"1\" w:author=\"A\"><w:trPr><w:tblHeader w:val=\"false\"/></w:trPr></w:trPrChange></w:trPr>"
        + "<w:tc><w:p><w:r><w:t>Region</w:t></w:r></w:p></w:tc></w:tr>"
        + "<w:tr><w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc></w:tr>", false)]
    public void TextProjectionMatchesTheLeanReadForTableShapes(string tableContent, bool streams) {
        byte[] package = CreateTableDocument(tableContent);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddWordHandler().Build();

        OfficeDocumentReadResult lean = reader.ReadDocument(package, "table.docx",
            new ReaderOptions { Projection = ReaderProjection.TextAndMetadata, MaxTableRows = 1 });
        OfficeDocumentReadResult text = reader.ReadDocument(package, "table.docx",
            new ReaderOptions { Projection = ReaderProjection.Text, MaxTableRows = 1 });

        Assert.Equal(lean.GetText(), text.GetText());
        Assert.Equal(streams, text.CapabilitiesUsed.Contains(StreamedCapability));
    }

    private static byte[] CreateTableDocument(string tableContent) {
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            document.AddParagraph("Before");
            document.Save();
        }
        using (WordprocessingDocument package = WordprocessingDocument.Open(stream, true)) {
            W.Body body = package.MainDocumentPart!.Document!.Body!;
            var table = new W.Table("<w:tbl xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" + tableContent + "</w:tbl>");
            body.InsertAfter(table, body.Elements<W.Paragraph>().First());
            body.InsertAfter(new W.Paragraph(new W.Run(new W.Text("After"))), table);
        }
        return stream.ToArray();
    }

    // Paragraphs, list items, a hyperlink, a line break and tab, and a table with a repeated header row.
    private static byte[] CreateDocument(bool withHeader) {
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            document.AddParagraph("Quarterly review");
            WordParagraph mixed = document.AddParagraph("Owner:");
            mixed.AddText("\tPat").AddBreak();
            mixed.AddText("second line");
            document.AddParagraph("See ").AddHyperLink("the dashboard", new Uri("https://example.com/"));
            WordList list = document.AddList(WordListStyle.Bulleted);
            list.AddItem("First item");
            list.AddItem("Second item");
            WordTable table = document.AddTable(3, 2);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Region";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Total";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "North";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "1,200";
            table.Rows[2].Cells[0].Paragraphs[0].Text = "South";
            table.RepeatHeaderRowAtTheTopOfEachPage = true;
            document.AddParagraph("Closing note");
            if (withHeader) {
                document.PageSetup.Header.AddParagraph("Running header");
            }
            document.Save();
        }
        return stream.ToArray();
    }
}
