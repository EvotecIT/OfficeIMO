using OfficeIMO.Word.Pdf;
using OfficeIMO.Word;
using System.IO;
using System.IO.Compression;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void Test_WordDocument_SaveAsPdf_MultiLevelLists() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfListSample.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfListSample.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordList numbered = document.AddList(WordListStyle.Headings111);
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
        byte[] startPattern = Encoding.ASCII.GetBytes("stream\n");
        byte[] endPattern = Encoding.ASCII.GetBytes("\nendstream");
        int start = IndexOf(bytes, startPattern, 0);
        Assert.True(start >= 0, "stream marker not found");
        start += startPattern.Length;
        int end = IndexOf(bytes, endPattern, start);
        Assert.True(end >= 0, "endstream marker not found");
        int length = end - start;
        int deflateLength = length - 6;
        byte[] deflateData = new byte[deflateLength];
        Array.Copy(bytes, start + 2, deflateData, 0, deflateLength);
        using MemoryStream ms = new MemoryStream(deflateData);
        using DeflateStream ds = new DeflateStream(ms, CompressionMode.Decompress);
        using StreamReader reader = new StreamReader(ds, Encoding.GetEncoding("ISO-8859-1"));
        string content = reader.ReadToEnd();
        Assert.False(string.IsNullOrEmpty(content));
    }
}
