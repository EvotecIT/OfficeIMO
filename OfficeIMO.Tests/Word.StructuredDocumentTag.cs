using OfficeIMO.Word;
using Xunit;
using System.IO;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_AddingStructuredDocumentTag() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithContentControl.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var sdt = document.AddStructuredDocumentTag("Hello world", "Alias1");

                Assert.True(document.StructuredDocumentTags.Count == 1);
                Assert.True(document.ParagraphsStructuredDocumentTags.Count == 1);
                Assert.Equal("Hello world", sdt.Text);

                document.Save(false);

                Assert.False(HasUnexpectedElements(document), "Document has unexpected elements. Order of elements matters!");
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.True(document.StructuredDocumentTags.Count == 1);
                Assert.Equal("Hello world", document.StructuredDocumentTags[0].Text);

                document.StructuredDocumentTags[0].Text = "Changed";
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal("Changed", document.StructuredDocumentTags[0].Text);
            }
        }

        [Fact]
        public void Test_StructuredDocumentTagWithTag() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithContentControlTag.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var sdt = document.AddStructuredDocumentTag("Hello", "Alias1", "Tag1");

                Assert.Equal("Tag1", sdt.Tag);
                document.Save(false);
                Assert.False(HasUnexpectedElements(document), "Document has unexpected elements. Order of elements matters!");
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var loaded = document.GetStructuredDocumentTagByTag("Tag1");
                Assert.NotNull(loaded);
                Assert.Equal("Hello", loaded.Text);

                loaded.Text = "Updated";
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal("Updated", document.StructuredDocumentTags[0].Text);
                Assert.Equal("Tag1", document.StructuredDocumentTags[0].Tag);
            }
        }
    }
}
