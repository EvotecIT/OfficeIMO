using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_HyperlinkWithWhitespaceOnlyRunKeepsTheSpaceAndBothTargets(bool inTableCell) {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = inTableCell
            ? source.AddTable(1, 1, WordTableStyle.TableNormal).Rows[0].Cells[0].Paragraphs[0]
            : source.AddParagraph();
        paragraph.AddHyperLink("LINKSTART", new Uri("https://www.example.com/"));
        W.Hyperlink hyperlink = paragraph._paragraph.Elements<W.Hyperlink>().Last();
        hyperlink.Append(
            new W.Run(new W.Text(" ") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }),
            new W.Run(new W.Text("LINKEND")));

        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        using var output = new MemoryStream();
        document.SaveAsPdf(output, new WordPdfExportOptions());

        byte[] bytes = output.ToArray();
        PdfDocumentInfo navigation = PdfInspector.Inspect(bytes);
        Assert.NotEmpty(navigation.LinkAnnotations);
        Assert.All(navigation.LinkAnnotations, link => Assert.Equal("https://www.example.com/", link.Uri));
        using var pdf = PdfPigDocument.Open(bytes);
        var page = pdf.GetPage(1);
        var start = FindRichNoteToken(page.Letters.ToArray(), "LINKSTART");
        var end = FindRichNoteToken(page.Letters.ToArray(), "LINKEND");
        Assert.True(end[0].StartBaseLine.X - start[start.Length - 1].EndBaseLine.X > 1D);
        foreach (var token in new[] { start, end }) {
            double x = (token[0].StartBaseLine.X + token[token.Length - 1].EndBaseLine.X) / 2D;
            double y = token[0].StartBaseLine.Y;
            Assert.Contains(navigation.LinkAnnotations, link => link.X1 <= x && x <= link.X2 && link.Y1 <= y && y <= link.Y2);
        }
    }
}
