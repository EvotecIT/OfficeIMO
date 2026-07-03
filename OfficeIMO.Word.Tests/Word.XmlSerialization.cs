using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_ParagraphXmlRoundTrip() {
            string filePath = Path.Combine(_directoryWithFiles, "ParagraphXmlRoundTrip.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Example text");
                string xml = paragraph.ToXml();

                var cloned = document.AddParagraphFromXml(xml);

                Assert.Equal(paragraph.ToXml(), cloned.ToXml());
                Assert.Equal(2, document.Paragraphs.Count);
                document.Save(false);
            }
        }
    }
}
