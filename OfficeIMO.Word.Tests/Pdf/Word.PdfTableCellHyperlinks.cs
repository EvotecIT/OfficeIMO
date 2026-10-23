using DocumentFormat.OpenXml;
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
    public void SaveAsPdf_TableCellKeepsExternalAndAnchorLinksLocalToTheirText(bool secondExternal) {
        using WordDocument source = CreateFootnoteDiagnosticDocument();
        source.AddParagraph("LINKTARGET").AddBookmark("CellLinkTarget");
        WordTableCell cell = source.AddTable(1, 1, WordTableStyle.TableNormal).Rows[0].Cells[0];
        cell.Paragraphs[0].Text = "PLAINCELLTEXT";
        cell.AddParagraph(" ").AddHyperLink("FIRSTLINK", new Uri("https://example.test/first"), tooltip: "First link tooltip");
        cell._tableCell.Append(new W.Paragraph(new W.Hyperlink(RichNoteRun("ANCHORLINK")) { Anchor = "CellLinkTarget" }));
        if (secondExternal) cell.AddParagraph(" ").AddHyperLink("SECONDLINK", new Uri("https://example.test/second"), tooltip: "Second link tooltip");
        Assert.Empty(source.ValidateDocument());

        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        var main = document._wordprocessingDocument.MainDocumentPart!;
        string original = main.Document.OuterXml;
        PdfDocumentConversionResult result = document.ToPdfDocumentResult(FootnoteDiagnosticOptions());
        for (int serialization = 0; serialization < 2; serialization++) {
            byte[] bytes = result.ToBytes();
            using var pdf = PdfPigDocument.Open(bytes);
            PdfDocumentInfo navigation = PdfInspector.Inspect(bytes);
            Assert.Equal(secondExternal ? 3 : 2, navigation.LinkAnnotations.Count);
            PdfLinkAnnotation external = Assert.Single(navigation.LinkAnnotations, link => link.Uri == "https://example.test/first");
            Assert.Equal("First link tooltip", external.Contents);
            AssertTableCellLinkBounds(pdf, external, "FIRSTLINK");
            PdfLinkAnnotation anchor = Assert.Single(navigation.LinkAnnotations, link => link.DestinationName == "CellLinkTarget");
            AssertTableCellLinkBounds(pdf, anchor, "ANCHORLINK");
            Assert.Single(navigation.NamedDestinations, target => target.Name == "CellLinkTarget");
            if (secondExternal) {
                PdfLinkAnnotation second = Assert.Single(navigation.LinkAnnotations, link => link.Uri == "https://example.test/second");
                Assert.Equal("Second link tooltip", second.Contents);
                AssertTableCellLinkBounds(pdf, second, "SECONDLINK");
            }
            var page = pdf.GetPages().Single(item => item.Text.Contains("PLAINCELLTEXT", StringComparison.Ordinal));
            var plain = FindRichNoteToken(page.Letters.ToArray(), "PLAINCELLTEXT");
            double plainX = (plain[0].StartBaseLine.X + plain[plain.Length - 1].EndBaseLine.X) / 2D;
            double plainY = plain[0].StartBaseLine.Y;
            Assert.DoesNotContain(navigation.LinkAnnotations, link => link.PageNumber == page.Number &&
                link.X1 <= plainX && plainX <= link.X2 && link.Y1 <= plainY && plainY <= link.Y2);
            Assert.Equal(original, main.Document.OuterXml);
        }
    }

    private static void AssertTableCellLinkBounds(PdfPigDocument pdf, PdfLinkAnnotation link, string token) {
        int pageNumber = Assert.IsType<int>(link.PageNumber);
        var page = pdf.GetPage(pageNumber);
        var letters = FindRichNoteToken(page.Letters.ToArray(), token);
        Assert.True(link.X1 <= letters.Min(letter => letter.StartBaseLine.X) + .01D);
        Assert.True(link.X2 >= letters.Max(letter => letter.EndBaseLine.X) - .01D);
        Assert.All(letters, letter => Assert.InRange(letter.StartBaseLine.Y, link.Y1 - .01D, link.Y2 + .01D));
        Assert.True(link.X2 - link.X1 < 200D, "A run link must not cover the full table cell.");
    }
}
