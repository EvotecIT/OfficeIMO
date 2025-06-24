using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_RemovingPageBreakParagraph() {
            string filePath = Path.Combine(_directoryWithFiles, "RemovePageBreakParagraph.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Paragraph 1");
                document.AddPageBreak();
                document.AddParagraph("Paragraph 2");

                Assert.True(document.Paragraphs[1].IsPageBreak);

                document.Paragraphs[1].Remove();

                Assert.Equal(2, document.Paragraphs.Count);
                Assert.Empty(document.PageBreaks);

                document.Save(false);
            }
        }
    }
}
