using OfficeIMO.Word;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_StructuredDocumentTagInsideTable() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithTableContentControl.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var table = document.AddTable(1, 1);
                table.Rows[0].Cells[0].Paragraphs[0].AddStructuredDocumentTag("Cell", "AliasCell", "TagCell");
                var cb = table.Rows[0].Cells[0].Paragraphs[0].AddCheckBox(false, "AliasCheck", "TagCheck");
                document.Save(false);
                Assert.False(HasUnexpectedElements(document), "Document has unexpected elements. Order of elements matters!");
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var table = document.Tables[0];
                Assert.Equal(2, table.StructuredDocumentTags.Count);
                Assert.Single(table.CheckBoxes);
                Assert.Equal("TagCell", table.StructuredDocumentTags[0].Tag);
                Assert.Equal("TagCheck", table.CheckBoxes[0].Tag);
            }
        }
    }
}
