using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Tests related to inserting HTML fragments after a paragraph.
    /// </summary>
    public partial class Word {
        /// <summary>
        /// Adds an HTML fragment after an existing paragraph.
        /// </summary>
        [Theory]
        [InlineData("<!DOCTYPE html><html><head><style>p { color: blue; }</style></head><body><p>frag</p></body></html>")]
        [InlineData("<!-- <html> is only a comment --><script>const root = '</scripture><html>';</script><p>frag</p>")]
        public void Test_AddEmbeddedFragmentAfter(string html) {
            string filePath = Path.Combine(_directoryWithFiles, "FragmentAfter.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var p1 = document.AddParagraph("Before");
                document.AddParagraph("After");
                document.AddEmbeddedFragmentAfter(p1, html);

                Assert.Single(document.EmbeddedDocuments);
                var body = document._document.Body!;
                Assert.IsType<Paragraph>(body.ChildElements[0]);
                Assert.IsType<AltChunk>(body.ChildElements[1]);
                Assert.IsType<Paragraph>(body.ChildElements[2]);
                Assert.IsType<SectionProperties>(body.ChildElements[3]);
                document.Save();
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                string stored = Assert.IsType<string>(Assert.Single(document.EmbeddedDocuments).GetHtml());
                if (html.StartsWith("<!DOCTYPE")) {
                    Assert.Equal(html, stored);
                } else {
                    Assert.StartsWith("<html>", stored);
                    Assert.Contains(html, stored);
                    Assert.EndsWith("</body></html>", stored);
                }
            }
        }
    }
}
