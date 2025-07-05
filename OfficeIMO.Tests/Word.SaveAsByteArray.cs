using System.IO;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_SaveAsByteArray() {
            using var document = WordDocument.Create();
            document.AddParagraph("Hello bytes");
            byte[] data = document.SaveAsByteArray();

            using var ms = new MemoryStream(data);
            using var openXml = WordprocessingDocument.Open(ms, false);
            Assert.NotNull(openXml.MainDocumentPart);
            ms.Position = 0;
            using var loaded = WordDocument.Load(ms);
            Assert.Equal("Hello bytes", loaded.Paragraphs[0].Text);
        }

        [Fact]
        public void Test_SaveAsMemoryStream() {
            using var document = WordDocument.Create();
            document.AddParagraph("Hello memory");

            using MemoryStream ms = document.SaveAsMemoryStream();
            using var openXml = WordprocessingDocument.Open(ms, false);
            Assert.NotNull(openXml.MainDocumentPart);
            ms.Position = 0;
            using var loaded = WordDocument.Load(ms);
            Assert.Equal("Hello memory", loaded.Paragraphs[0].Text);
        }

        [Fact]
        public void Test_SaveAsStream() {
            using var document = WordDocument.Create();
            document.AddParagraph("Hello stream");

            using var ms = new MemoryStream();
            using var clone = document.SaveAs(ms);

            Assert.Equal(string.Empty, document.FilePath);
            Assert.Null(clone.FilePath);
            Assert.Single(clone.Paragraphs);
            Assert.Equal("Hello stream", clone.Paragraphs[0].Text);
        }
    }
}
