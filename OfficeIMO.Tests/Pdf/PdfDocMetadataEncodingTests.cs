using System;
using System.IO;
using OfficeIMO.Pdf;
using UglyToad.PdfPig;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public class PdfDocMetadataEncodingTests {
        [Fact]
        public void PdfDoc_Metadata_Uses_Latin1_Bytes() {
            const string title = "Caf\u00e9 num\u00e9ro";
            const string author = "M\u00fcller & Co.";
            const string subject = "Line1\nLine2\tTabbed";
            const string keywords = "jalape\u00f1o, fa\u00e7ade";

            var doc = PdfDoc.Create();
            doc.Meta(title: title, author: author, subject: subject, keywords: keywords);
            doc.Compose(c =>
                c.Page(page =>
                    page.Content(content =>
                        content.Column(col =>
                            col.Item().Paragraph(p => p.Text("Hello metadata"))))));

            byte[] pdfBytes = doc.ToBytes();
            Assert.NotEmpty(pdfBytes);

            AssertContains(pdfBytes, "/Title (" + title + ")", "Title should be written with Latin-1 bytes.");
            AssertContains(pdfBytes, "/Author (" + author + ")", "Author should be written with Latin-1 bytes.");
            AssertContains(pdfBytes, "/Subject (Line1\\nLine2\\tTabbed)", "Subject should escape control characters using text escapes.");
            AssertContains(pdfBytes, "/Keywords (" + keywords + ")", "Keywords should retain extended characters.");

            using var pdf = PdfDocument.Open(new MemoryStream(pdfBytes));
            var info = pdf.Information;
            Assert.Equal(title, info.Title);
            Assert.Equal(author, info.Author);
            Assert.Equal(subject, info.Subject);
            Assert.Equal(keywords, info.Keywords);
        }

        private static void AssertContains(byte[] haystack, string text, string message) {
            var pattern = PdfEncoding.Latin1GetBytes(text);
            Assert.True(ContainsSequence(haystack, pattern), message);
        }

        private static bool ContainsSequence(byte[] haystack, byte[] needle) {
            if (needle.Length == 0) return true;
            for (int i = 0; i <= haystack.Length - needle.Length; i++) {
                if (haystack[i] != needle[0]) continue;
                int j = 1;
                for (; j < needle.Length; j++) {
                    if (haystack[i + j] != needle[j]) break;
                }
                if (j == needle.Length) return true;
            }
            return false;
        }
    }
}
