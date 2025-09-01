using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System.IO;
using System.Linq;
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
        public void CleanupDocument_MergesRunsWithDifferentAttributeOrder() {
            string filePath = Path.Combine(_directoryWithFiles, "CleanupRunsAttributes.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordParagraph p = document.AddParagraph();

                var color1 = new Color();
                color1.SetAttribute(new OpenXmlAttribute("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "FF0000"));
                color1.SetAttribute(new OpenXmlAttribute("w", "themeColor", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "accent1"));

                var color2 = new Color();
                color2.SetAttribute(new OpenXmlAttribute("w", "themeColor", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "accent1"));
                color2.SetAttribute(new OpenXmlAttribute("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "FF0000"));

                var run1 = new Run(new RunProperties(color1), new Text("A"));
                var run2 = new Run(new RunProperties(color2), new Text("B"));

                p._paragraph.Append(run1, run2);

                Assert.Equal(2, p._paragraph.Elements<Run>().Count());

                document.CleanupDocument(DocumentCleanupOptions.MergeIdenticalRuns);

                Assert.Single(p._paragraph.Elements<Run>());
                Assert.Equal("AB", p._paragraph.InnerText);
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
                Assert.NotNull(run._run);
                Assert.NotNull(run._run!.RunProperties);
                document.CleanupDocument();
                Assert.NotNull(run._paragraph);
                var firstRun = run._paragraph!.Elements<Run>().First();
                Assert.Null(firstRun.RunProperties);
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