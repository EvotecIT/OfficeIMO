using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_ReplaceTextWithHtmlFragment_Simple() {
            string filePath = Path.Combine(_directoryWithFiles, "ReplaceHtmlSimple.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Hello replaceTarget world.");
                var count = document.ReplaceTextWithHtmlFragment("replaceTarget", "<html><p>Injected</p></html>");
                Assert.Equal(1, count);
                Assert.Single(document.EmbeddedDocuments);
                Assert.DoesNotContain("replaceTarget", document.Paragraphs[0].Text);
                var body = document._document.Body;
                Assert.IsType<SectionProperties>(body.ChildElements[0]);
                Assert.IsType<Paragraph>(body.ChildElements[1]);
                Assert.IsType<AltChunk>(body.ChildElements[2]);
                document.Save();
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Single(document.EmbeddedDocuments);
            }
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
                var body = document._document.Body;
                Assert.IsType<SectionProperties>(body.ChildElements[0]);
                Assert.IsType<Paragraph>(body.ChildElements[1]);
                Assert.IsType<AltChunk>(body.ChildElements[2]);
                Assert.IsType<Paragraph>(body.ChildElements[3]);
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
