using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
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

        [Fact]
        public void PdfDoc_Metadata_With_Emoji_Uses_Utf16_Hex() {
            const string title = "Emoji 😀 Title";
            const string author = "漢字カタカナ";
            const string subject = "Lock 🔒 Subject";
            const string keywords = "mix 😀 漢字";

            var doc = PdfDoc.Create();
            doc.Meta(title: title, author: author, subject: subject, keywords: keywords);
            doc.Compose(c =>
                c.Page(page =>
                    page.Content(content =>
                        content.Column(col =>
                            col.Item().Paragraph(p => p.Text("Unicode metadata"))))));

            byte[] pdfBytes = doc.ToBytes();
            Assert.NotEmpty(pdfBytes);

            AssertContains(pdfBytes, "/Title <" + Utf16Hex(title) + ">", "Title should be encoded as UTF-16BE hex when non-Latin1 characters are present.");
            AssertContains(pdfBytes, "/Author <" + Utf16Hex(author) + ">", "Author should be encoded as UTF-16BE hex when non-Latin1 characters are present.");
            AssertContains(pdfBytes, "/Subject <" + Utf16Hex(subject) + ">", "Subject should be encoded as UTF-16BE hex when non-Latin1 characters are present.");
            AssertContains(pdfBytes, "/Keywords <" + Utf16Hex(keywords) + ">", "Keywords should be encoded as UTF-16BE hex when non-Latin1 characters are present.");

            using var pdf = PdfDocument.Open(new MemoryStream(pdfBytes));
            var info = pdf.Information;
            Assert.Equal(title, info.Title);
            Assert.Equal(author, info.Author);
            Assert.Equal(subject, info.Subject);
            Assert.Equal(keywords, info.Keywords);
        }

        [Fact]
        public void Latin1GetBytes_Throws_For_NonLatin1_Input() {
            Assert.Throws<ArgumentException>(() => PdfEncoding.Latin1GetBytes("😀"));
            Assert.Throws<ArgumentException>(() => PdfEncoding.Latin1GetBytes("漢"));
        }

        [Fact]
        public void Footer_Text_With_NonLatin1_Characters_Falls_Back_To_WinAnsi() {
            var options = new PdfOptions {
                ShowPageNumbers = true,
                FooterSegments = new List<FooterSegment> {
                    new FooterSegment(FooterSegmentKind.Text, "😀 footer")
                }
            };

            var doc = PdfDoc.Create(options);
            doc.Compose(c =>
                c.Page(page =>
                    page.Content(content =>
                        content.Column(col =>
                            col.Item().Paragraph(p => p.Text("Body"))))));

            byte[] pdfBytes = doc.ToBytes();
            Assert.NotEmpty(pdfBytes);

            AssertContains(pdfBytes, "<3F3F20666F6F746572> Tj", "Footer text should be encoded with WinAnsi hex, replacing surrogate pairs with '?' bytes to keep a valid stream.");
        }

        private static void AssertContains(byte[] haystack, string text, string message) {
            var pattern = PdfEncoding.Latin1GetBytes(text);
            Assert.True(ContainsSequence(haystack, pattern), message);
        }

        private static string Utf16Hex(string text) {
            var bytes = Encoding.BigEndianUnicode.GetBytes(text);
            var sb = new StringBuilder((bytes.Length + 2) * 2);
            sb.Append("FEFF");
            for (int i = 0; i < bytes.Length; i++) {
                sb.Append(bytes[i].ToString("X2", System.Globalization.CultureInfo.InvariantCulture));
            }
            return sb.ToString();
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
