using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void CleanupDocument_MergesRunsWithSameFormatting() {
            string filePath = Path.Combine(_directoryWithFiles, "CleanupRuns.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordParagraph p = document.AddParagraph("Hello");
                p.SetBold();
                p.AddText(" World").SetBold();
                Assert.Equal(2, p._paragraph.Elements<Run>().Count());
                document.CleanupDocument();
                Assert.Single(p._paragraph.Elements<Run>());
                Assert.Equal("Hello World", p.Text);
                document.Save(false);
            }
        }

        [Fact]
        public void CleanupDocument_RemovesEmptyRuns() {
            string filePath = Path.Combine(_directoryWithFiles, "CleanupEmptyRuns.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordParagraph p = document.AddParagraph("Hello");
                p._paragraph.AppendChild(new Run());
                Assert.Equal(2, p._paragraph.Elements<Run>().Count());
                document.CleanupDocument();
                Assert.Single(p._paragraph.Elements<Run>());
                document.Save(false);
            }
        }

        [Fact]
        public void CleanupDocument_RemovesRedundantRunProperties() {
            string filePath = Path.Combine(_directoryWithFiles, "CleanupRunProperties.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordParagraph run = document.AddParagraph().AddText("Test");
                run.SetBold();
                run.SetBold(false);
                Assert.NotNull(run._run.RunProperties);
                document.CleanupDocument();
                Assert.Null(run._paragraph.Elements<Run>().First().RunProperties);
                document.Save(false);
            }
        }

        [Fact]
        public void CleanupDocument_RemovesEmptyParagraphs() {
            string filePath = Path.Combine(_directoryWithFiles, "CleanupEmptyParagraphs.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Hello");
                document.AddParagraph();
                Assert.Equal(2, document.Paragraphs.Count);
                document.CleanupDocument();
                Assert.Single(document.Paragraphs);
                Assert.Equal("Hello", document.Paragraphs.First().Text);
                document.Save(false);
            }
        }
    }
}
