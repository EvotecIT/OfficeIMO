using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using SixLabors.ImageSharp;
using System.IO;
using System.IO.Compression;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void Test_WordDocument_SaveAsPdf_Shapes() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfShapesSample.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfShapesSample.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            var paragraph = document.AddParagraph();
            paragraph.AddShape(ShapeType.Rectangle, 80, 40, Color.Aqua, Color.Black, 1);
            WordShape.AddLine(paragraph, 0, 50, 80, 50, Color.Red, 2);
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
        if (length > 6) {
            int deflateLength = length - 6;
            byte[] deflateData = new byte[deflateLength];
            System.Array.Copy(bytes, start + 2, deflateData, 0, deflateLength);
            using MemoryStream ms = new MemoryStream(deflateData);
            using DeflateStream ds = new DeflateStream(ms, CompressionMode.Decompress);
            using StreamReader reader = new StreamReader(ds, Encoding.GetEncoding("ISO-8859-1"));
            string content = reader.ReadToEnd();
            Assert.False(string.IsNullOrEmpty(content));
        }
    }
}
