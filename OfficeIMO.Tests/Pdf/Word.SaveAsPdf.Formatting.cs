using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System.IO;
using System.IO.Compression;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_WordDocument_SaveAsPdf_ParagraphFormatting() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfParagraphFormatting.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfParagraphFormatting.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordParagraph paragraph = document.AddParagraph("Formatted Text");
                paragraph.FontSize = 20;
                paragraph.Strike = true;
                paragraph.Highlight = HighlightColorValues.Yellow;
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

            Assert.Contains("/F4 20 Tf", content);
            Assert.Contains("1 1 0 rg", content);
            Assert.Contains("13.795", content);
        }
    }
}

