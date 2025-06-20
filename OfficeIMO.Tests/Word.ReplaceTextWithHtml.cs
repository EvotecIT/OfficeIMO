using System.IO;
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
                document.Save();
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Single(document.EmbeddedDocuments);
            }
        }
    }
}
