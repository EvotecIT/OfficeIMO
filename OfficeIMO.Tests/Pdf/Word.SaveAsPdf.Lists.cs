using OfficeIMO.Word.Pdf;
using OfficeIMO.Word;
using System.IO;
using System.Linq;
using System.Text;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void Test_WordDocument_SaveAsPdf_MultiLevelLists() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfListSample.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfListSample.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordList numbered = document.AddList(WordListStyle.Numbered);
            numbered.AddItem("Numbered 1");
            numbered.AddItem("Numbered 2");
            numbered.AddItem("Numbered 2.1", 1);
            numbered.AddItem("Numbered 3");

            WordList bullet = document.AddList(WordListStyle.Bulleted);
            bullet.AddItem("Bullet 1");
            bullet.AddItem("Bullet 2");
            bullet.AddItem("Bullet 2.1", 1);
            bullet.AddItem("Bullet 3");

            document.Save();
            document.SaveAsPdf(pdfPath);
        }

        Assert.True(File.Exists(pdfPath));

        byte[] bytes = File.ReadAllBytes(pdfPath);
        string content = ReadFirstPdfStreamContent(bytes);
        Assert.False(string.IsNullOrEmpty(content));
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Charts_In_List_Paragraphs() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfListChart.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfListChart.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordList list = document.AddList(WordListStyle.Numbered);
            WordParagraph item = list.AddItem("List item with chart");
            WordChart chart = item.AddChart("List Chart", false, 240, 160);
            chart.AddPie("Passed", 2);
            chart.AddPie("Failed", 1);
            document.AddParagraph("After list chart");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        Assert.True(File.Exists(pdfPath));
        using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
            string text = string.Concat(pdf.GetPages().Select(page => page.Text));
            Assert.Contains("After list chart", text);
        }

        string content = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Assert.Contains("0.122 0.306 0.475 rg", content);
    }
}
