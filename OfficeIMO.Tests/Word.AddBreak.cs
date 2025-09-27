using System.IO;
using System.Linq;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void DocumentAddBreakIncreasesParagraphCount() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentAddBreakIncreasesParagraphCount.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Paragraph 1");
                int paragraphCount = document.Paragraphs.Count;
                document.AddBreak();
                Assert.Equal(paragraphCount + 1, document.Paragraphs.Count);
                Assert.Equal(paragraphCount + 1, document.Sections[0].Paragraphs.Count);
                Assert.Single(document.Breaks);
                document.Save(false);
            }
        }

        [Fact]
        public void SectionParagraphsIncludeBreakAfterDocumentAddBreak() {
            string filePath = Path.Combine(_directoryWithFiles, "SectionParagraphsIncludeBreakAfterDocumentAddBreak.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var section = document.Sections[0];
                int initialParagraphCount = section.Paragraphs.Count;

                document.AddBreak();

                var sectionParagraphs = section.Paragraphs;
                Assert.Equal(initialParagraphCount + 1, sectionParagraphs.Count);
                Assert.True(sectionParagraphs.Last().IsPageBreak);
                document.Save(false);
            }
        }
    }
}
