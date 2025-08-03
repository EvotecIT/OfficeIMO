using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void Test_WordDocument_SaveAsPdf() {
        var docPath = Path.Combine(_directoryWithFiles, "PdfSample.docx");
        var pdfPath = Path.Combine(_directoryWithFiles, "PdfSample.pdf");

        using (var document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            document.Header.Default.AddParagraph("Sample Header");
            document.Footer.Default.AddParagraph("Sample Footer");

            var heading = document.AddParagraph("Heading One");
            heading.Style = WordParagraphStyles.Heading1;

            var formatted = document.AddParagraph("Centered Bold Italic Underlined");
            formatted.Bold = true;
            formatted.Italic = true;
            formatted.Underline = UnderlineValues.Single;
            formatted.ParagraphAlignment = JustificationValues.Center;

            var list = document.AddList(WordListStyle.ArticleSections);
            list.AddItem("Numbered Item 1");
            list.AddItem("Numbered Item 2");

            var table = document.AddTable(2, 2);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "A1";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "B1";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "A2";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "B2";

            var imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
            document.AddParagraph().AddImage(imagePath, 50, 50);

            document.Save();
            document.SaveAsPdf(pdfPath);
        }

        Assert.True(File.Exists(pdfPath));
    }
}
