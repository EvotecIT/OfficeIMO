using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
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
        public void Test_SaveAsMemoryStream_RunsCompatibilityFixer() {
            using var document = WordDocument.Create();
            document.AddParagraph("Memory compatibility");

            using MemoryStream ms = document.SaveAsMemoryStream();
            Assert.True(ms.CanRead);
            Assert.Equal(0, ms.Position);

            using (var package = Package.Open(ms, FileMode.Open, FileAccess.Read)) {
                var relsUri = new Uri("/word/_rels/document.xml.rels", UriKind.Relative);
                Assert.True(package.PartExists(relsUri));
                var part = package.GetPart(relsUri);
                using var relStream = part.GetStream(FileMode.Open, FileAccess.Read);
                var rels = XDocument.Load(relStream);
                var targets = rels.Root?.Elements().Select(e => e.Attribute("Target")?.Value).Where(v => !string.IsNullOrEmpty(v)) ?? Enumerable.Empty<string>();
                Assert.All(targets, target => Assert.False(target!.StartsWith("/word/", StringComparison.Ordinal), $"Relationship target '{target}' should not start with '/word/'."));
            }

            ms.Position = 0;
            using var reloaded = WordDocument.Load(ms);
            var paragraph = Assert.Single(reloaded.Paragraphs);
            Assert.Equal("Memory compatibility", paragraph.Text);
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
