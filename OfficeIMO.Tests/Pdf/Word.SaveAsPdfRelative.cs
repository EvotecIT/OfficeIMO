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
                document.AddParagraph("Hello PDF");
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