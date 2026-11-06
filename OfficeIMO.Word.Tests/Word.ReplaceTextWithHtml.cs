using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void ReplaceTextWithHtmlFragment_PreservesTextAndPlacementForMatchesInOneRun() {
            using var output = new MemoryStream();
            using (WordDocument document = WordDocument.Create()) {
                document.AddParagraph("first TOKEN middle TTOKEN tail");
                int count = document.ReplaceTextWithHtmlFragment("TOKEN", "<html><p>Injected</p></html>");

                Assert.Equal(2, count);
                document.Save(output);
            }

            output.Position = 0;
            using WordDocument reopened = WordDocument.Load(output, new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
            Assert.Equal("first  middle T tail", Assert.Single(reopened.Paragraphs).Text);
            Assert.Equal(2, reopened.EmbeddedDocuments.Count);
            var body = reopened._document.Body!;
            Assert.IsType<Paragraph>(body.ChildElements[0]);
            Assert.IsType<AltChunk>(body.ChildElements[1]);
            Assert.IsType<AltChunk>(body.ChildElements[2]);
            Assert.IsType<SectionProperties>(body.ChildElements[3]);
        }

        [Fact]
        public void ReplaceTextWithHtmlFragment_CultureComparisonPreservesTextWithDifferentMatchWidth() {
            using WordDocument document = WordDocument.Create();
            const string text = "caf\u00e9 next";
            WordParagraph paragraph = document.AddParagraph(text);

            int count = document.ReplaceTextWithHtmlFragment("cafe\u0301", "<html><p>Injected</p></html>",
                stringComparison: StringComparison.InvariantCulture);

            Assert.Equal(0, count);
            Assert.Equal(text, paragraph.Text);
            Assert.Empty(document.EmbeddedDocuments);
        }

        [Fact]
        public void Test_ReplaceTextWithHtmlFragment_MultiParagraph() {
            string filePath = Path.Combine(_directoryWithFiles, "ReplaceHtmlMulti.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Intro start");
                document.AddParagraph("finish end");
                var count = document.ReplaceTextWithHtmlFragment("startfinish", "<html><p>Injected</p></html>");
                Assert.Equal(1, count);
                Assert.Single(document.EmbeddedDocuments);
                Assert.Equal("Intro ", document.Paragraphs[0].Text);
                Assert.Equal(" end", document.Paragraphs[1].Text);
                var body = document._document.Body!;
                Assert.IsType<Paragraph>(body.ChildElements[0]);
                Assert.IsType<AltChunk>(body.ChildElements[1]);
                Assert.IsType<Paragraph>(body.ChildElements[2]);
                Assert.IsType<SectionProperties>(body.ChildElements[3]);
                document.Save();
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Single(document.EmbeddedDocuments);
            }
        }

        [Fact]
        public void Test_ReplaceTextWithHtmlFragment_BoundaryStart() {
            string filePath = Path.Combine(_directoryWithFiles, "ReplaceHtmlBoundaryStart.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("target middle");
                var count = document.ReplaceTextWithHtmlFragment("target", "<html><p>Injected</p></html>");
                Assert.Equal(1, count);
                Assert.Single(document.EmbeddedDocuments);
                Assert.Equal(" middle", document.Paragraphs[0].Text);
                document.Save();
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Single(document.EmbeddedDocuments);
            }
        }

        [Fact]
        public void Test_ReplaceTextWithHtmlFragment_BoundaryEnd() {
            string filePath = Path.Combine(_directoryWithFiles, "ReplaceHtmlBoundaryEnd.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Some text");
                document.AddParagraph("final target");
                var count = document.ReplaceTextWithHtmlFragment("target", "<html><p>Injected</p></html>");
                Assert.Equal(1, count);
                Assert.Single(document.EmbeddedDocuments);
                Assert.Equal("final ", document.Paragraphs[1].Text);
                document.Save();
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Single(document.EmbeddedDocuments);
            }
        }
    }
}
