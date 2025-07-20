using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CrossReferenceBookmark() {
            string filePath = Path.Combine(_directoryWithFiles, "CrossRefBookmark.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Start").AddBookmark("B1");
                document.AddParagraph("See ").AddCrossReference("B1", WordCrossReferenceType.Bookmark).AddText(".");
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Single(document.Fields);
                Assert.Equal(WordFieldType.Ref, document.Fields[0].FieldType);
                Assert.Equal(new[] { "B1" }, document.Fields[0].FieldInstructions);
            }
        }

        [Fact]
        public void Test_CrossReferenceHeading() {
            string filePath = Path.Combine(_directoryWithFiles, "CrossRefHeading.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var heading = document.AddParagraph("Header");
                heading.Style = WordParagraphStyles.Heading1;
                heading.AddBookmark("H1");

                document.AddParagraph("See heading ").AddCrossReference("H1", WordCrossReferenceType.Heading).AddText(".");
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Single(document.Fields);
                Assert.Equal("H1", document.Fields[0].FieldInstructions[0]);
            }
        }
    }
}
