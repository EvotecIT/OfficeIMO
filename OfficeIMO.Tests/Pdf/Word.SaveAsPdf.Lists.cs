using OfficeIMO.Word.Pdf;
using OfficeIMO.Word;
using System.IO;
using System.Text;
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
}
