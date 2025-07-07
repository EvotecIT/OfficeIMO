using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;
using System.IO;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Tests related to inserting HTML fragments after a paragraph.
    /// </summary>
    public partial class Word {
        /// <summary>
        /// Adds an HTML fragment after an existing paragraph.
        /// </summary>
        [Fact]
        public void Test_AddEmbeddedFragmentAfter() {
            string filePath = Path.Combine(_directoryWithFiles, "FragmentAfter.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var p1 = document.AddParagraph("Before");
                document.AddParagraph("After");
                document.AddEmbeddedFragmentAfter(p1, "<html><p>frag</p></html>");

                Assert.Single(document.EmbeddedDocuments);
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
    }
}
