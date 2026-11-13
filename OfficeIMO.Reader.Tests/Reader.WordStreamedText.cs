using OfficeIMO.Reader;
using OfficeIMO.Reader.Word;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderWordStreamedTextTests {
    private const string StreamedCapability = "officeimo.reader.word.streamed-text";

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StreamedTextMatchesTheModelProjectionAndDeclinesRunningHeaders(bool withHeader) {
        byte[] package = CreateDocument(withHeader);
        var options = new ReaderOptions { Projection = ReaderProjection.TextAndMetadata };

        OfficeDocumentReadResult model = new OfficeDocumentReaderBuilder().AddWordHandler().Build()
            .ReadDocument(package, "streamed.docx", options);
        OfficeDocumentReadResult streamed = new OfficeDocumentReaderBuilder()
            .AddWordHandler(new ReaderWordOptions { StreamTextWhenPossible = true }).Build()
            .ReadDocument(package, "streamed.docx", options);

        Assert.Equal(model.GetText(), streamed.GetText());
        Assert.Equal(!withHeader, streamed.CapabilitiesUsed.Contains(StreamedCapability));
        Assert.DoesNotContain(StreamedCapability, model.CapabilitiesUsed);
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
