using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Contains advanced content control tests.
    /// </summary>
    public partial class Word {
        [Fact]
        public void Test_AdvancedContentControls() {
            string folderPath = _directoryWithFiles;
            string filePath = Path.Combine(folderPath, "DocumentAdvancedContentControls.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var para1 = document.AddParagraph("Control 1:");
                para1.AddStructuredDocumentTag("First", "Alias1", "Tag1");

                var para2 = document.AddParagraph("Control 2:");
                para2.AddStructuredDocumentTag("Second", "Alias2", "Tag2");

                var para3 = document.AddParagraph("Control 3:");
                para3.AddStructuredDocumentTag("Third", "Alias3", "Tag3");

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var aliasControl = document.GetStructuredDocumentTagByAlias("Alias2");
                Assert.NotNull(aliasControl);
                aliasControl.Text = "Changed";

                var tagControl = document.GetStructuredDocumentTagByTag("Tag3");
                Assert.NotNull(tagControl);
                tagControl.Text = "Modified";
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var aliasControl = document.GetStructuredDocumentTagByAlias("Alias2");
                Assert.NotNull(aliasControl);
                Assert.Equal("Changed", aliasControl!.Text);
                var tagControl = document.GetStructuredDocumentTagByTag("Tag3");
                Assert.NotNull(tagControl);
                Assert.Equal("Modified", tagControl!.Text);
            }
        }
    }
}
