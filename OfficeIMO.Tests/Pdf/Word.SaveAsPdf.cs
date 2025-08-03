using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using QuestPDF.Helpers;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
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

    [Theory]
    [InlineData(PdfPageOrientation.Portrait)]
    [InlineData(PdfPageOrientation.Landscape)]
    public void Test_WordDocument_SaveAsPdf_PageOrientation(PdfPageOrientation orientation) {
        string docPath = Path.Combine(_directoryWithFiles, $"PdfOrientation{orientation}.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, $"PdfOrientation{orientation}.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Hello World");
            document.Save();

            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                Orientation = orientation,
                PageSize = PageSizes.A4
            });
        }

        Assert.True(File.Exists(pdfPath));

        string pdfContent = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Match mediaBox = Regex.Match(pdfContent, @"/MediaBox\s*\[\s*0\s+0\s+(?<w>[0-9\.]+)\s+(?<h>[0-9\.]+)\s*\]");
        Assert.True(mediaBox.Success, "MediaBox not found");
        double width = double.Parse(mediaBox.Groups["w"].Value, CultureInfo.InvariantCulture);
        double height = double.Parse(mediaBox.Groups["h"].Value, CultureInfo.InvariantCulture);
        if (orientation == PdfPageOrientation.Landscape) {
            Assert.True(width > height);
        } else {
            Assert.True(height > width);
        }
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_CustomPageSize() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfCustomSize.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfCustomSize.pdf");
        QuestPDF.Helpers.PageSize size = new QuestPDF.Helpers.PageSize(300, 500);

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Hello World");
            document.Save();

            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                PageSize = size
            });
        }

        Assert.True(File.Exists(pdfPath));

        string pdfContent = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Match mediaBox = Regex.Match(pdfContent, @"/MediaBox\s*\[\s*0\s+0\s+(?<w>[0-9\.]+)\s+(?<h>[0-9\.]+)\s*\]");
        Assert.True(mediaBox.Success, "MediaBox not found");
        double width = double.Parse(mediaBox.Groups["w"].Value, CultureInfo.InvariantCulture);
        double height = double.Parse(mediaBox.Groups["h"].Value, CultureInfo.InvariantCulture);
        Assert.True(System.Math.Abs(width - size.Width) < 1);
        Assert.True(System.Math.Abs(height - size.Height) < 1);
    }
}