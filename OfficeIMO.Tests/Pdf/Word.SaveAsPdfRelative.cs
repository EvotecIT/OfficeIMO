using OfficeIMO.Pdf;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void Test_WordDocument_SaveAsPdf_RelativePath() {
        var docPath = Path.Combine(_directoryWithFiles, "RelativeSample.docx");
        var previous = Directory.GetCurrentDirectory();
        Directory.SetCurrentDirectory(_directoryWithFiles);
        try {
            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("First paragraph");
                document.AddParagraph("Second paragraph").Bold = true;

                WordList list = document.AddList(WordListStyle.Bulleted);
                list.AddItem("Bullet 1");
                list.AddItem("Bullet 2");

                WordTable table = document.AddTable(2, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "A1";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "B1";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "A2";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "B2";

                document.SaveAsPdf("output.pdf");
            }
        } finally {
            Directory.SetCurrentDirectory(previous);
        }

        var pdfPath = Path.Combine(_directoryWithFiles, "output.pdf");
        Assert.True(File.Exists(pdfPath));
        byte[] bytes = File.ReadAllBytes(pdfPath);
        Assert.True(bytes.Length > 100);
        string header = System.Text.Encoding.ASCII.GetString(bytes, 0, 4);
        Assert.Equal("%PDF", header);
        string trailer = System.Text.Encoding.ASCII.GetString(bytes, bytes.Length - 5, 5);
        Assert.Contains("EOF", trailer);
    }
}