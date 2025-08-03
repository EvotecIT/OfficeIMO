using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using System.Globalization;
using System.Text.RegularExpressions;
using Xunit;
using QuestPDF.Infrastructure;
using QuestPDF.Helpers;
using PageSize = QuestPDF.Helpers.PageSize;

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
    public void Test_WordDocument_SaveAsPdf_Landscape() {
        var docPath = Path.Combine(_directoryWithFiles, "PdfLandscape.docx");
        var pdfPath = Path.Combine(_directoryWithFiles, "PdfLandscape.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Landscape");
            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions { Orientation = PdfPageOrientation.Landscape });
        }

        (double width, double height) = GetPdfSize(pdfPath);
        Assert.True(width > height);
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_CustomSize() {
        var docPath = Path.Combine(_directoryWithFiles, "PdfCustom.docx");
        var pdfPath = Path.Combine(_directoryWithFiles, "PdfCustom.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Custom Size");
            document.Save();
            var options = new PdfSaveOptions { PageSize = new PageSize(300, 400, Unit.Point) };
            document.SaveAsPdf(pdfPath, options);
        }

        (double width, double height) = GetPdfSize(pdfPath);
        Assert.InRange(width, 299, 301);
        Assert.InRange(height, 399, 401);
    }

    private static (double Width, double Height) GetPdfSize(string path) {
        string content = File.ReadAllText(path);
        Match match = Regex.Match(content, @"/MediaBox\s*\[\s*0\s+0\s+(?<w>\d+(?:\.\d+)?)\s+(?<h>\d+(?:\.\d+)?)\s*\]");
        double width = double.Parse(match.Groups["w"].Value, CultureInfo.InvariantCulture);
        double height = double.Parse(match.Groups["h"].Value, CultureInfo.InvariantCulture);
        return (width, height);
    }
}