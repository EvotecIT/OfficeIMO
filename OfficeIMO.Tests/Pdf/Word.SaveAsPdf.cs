using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void Test_WordDocument_SaveAsPdf() {
        var docPath = Path.Combine(_directoryWithFiles, "PdfSample.docx");
        var pdfPath = Path.Combine(_directoryWithFiles, "PdfSample.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            document.Header.Default.AddParagraph("Sample Header");
            WordTable headerTable = document.Header.Default.AddTable(1, 1);
            headerTable.Rows[0].Cells[0].Paragraphs[0].Text = "H1";
            document.Footer.Default.AddParagraph("Sample Footer");
            WordTable footerTable = document.Footer.Default.AddTable(1, 1);
            footerTable.Rows[0].Cells[0].Paragraphs[0].Text = "F1";

            WordParagraph heading = document.AddParagraph("Heading One");
            heading.Style = WordParagraphStyles.Heading1;

            WordParagraph formatted = document.AddParagraph("Centered Bold Italic Underlined");
            formatted.Bold = true;
            formatted.Italic = true;
            formatted.Underline = UnderlineValues.Single;
            formatted.ParagraphAlignment = JustificationValues.Center;

            WordList list = document.AddList(WordListStyle.ArticleSections);
            list.AddItem("Numbered Item 1");
            list.AddItem("Numbered Item 2");

            WordTable table = document.AddTable(2, 2);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "A1";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "B1";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "A2";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "B2";
            WordTable nested = table.Rows[0].Cells[0].AddTable(1, 1);
            nested.Rows[0].Cells[0].Paragraphs[0].Text = "N1";

            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
            document.AddParagraph().AddImage(imagePath, 50, 50);

            document.Save();
            document.SaveAsPdf(pdfPath);
        }

        Assert.True(File.Exists(pdfPath));
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_ToMemoryStream() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfStreamSample.docx");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Hello World");
            document.Save();

            using (MemoryStream pdfStream = new MemoryStream()) {
                document.SaveAsPdf(pdfStream);
                Assert.True(pdfStream.Length > 0);
            }
        }
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_ToFileStream() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfFileStreamSample.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfFileStreamSample.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Hello World");
            document.Save();

            using (FileStream fileStream = File.Create(pdfPath)) {
                document.SaveAsPdf(fileStream);
            }
        }

        Assert.True(File.Exists(pdfPath));
    }
}