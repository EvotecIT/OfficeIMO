using OfficeIMO.Word;
using System.IO;
using System.Linq;
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

        [Fact]
        public void Test_TableCheckBoxRetrievalByAliasAndTag() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentTableCheckBox.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var table = document.AddTable(1, 2);
                table.Rows[0].Cells[0].Paragraphs[0].AddCheckBox(true, "Alias1", "Tag1");
                table.Rows[0].Cells[1].Paragraphs[0].AddCheckBox(false, "Alias2", "Tag2");
                document.Save(false);
                Assert.False(HasUnexpectedElements(document), "Document has unexpected elements. Order of elements matters!");
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var table = document.Tables[0];
                Assert.Equal(2, table.CheckBoxes.Count);

                var cb1 = table.CheckBoxes.First(cb => cb.Tag == "Tag1");
                Assert.Equal("Alias1", cb1.Alias);
                Assert.True(cb1.IsChecked);

                var cb2 = table.CheckBoxes.First(cb => cb.Alias == "Alias2");
                Assert.False(cb2.IsChecked);
                Assert.Equal("Tag2", cb2.Tag);
            }
        }

        [Fact]
        public void Test_TableMultipleControlsCounts() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentTableMultipleControls.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var table = document.AddTable(2, 2);
                table.Rows[0].Cells[0].Paragraphs[0].AddStructuredDocumentTag("One", "AliasSdt1", "TagSdt1");
                table.Rows[0].Cells[1].Paragraphs[0].AddCheckBox(false, "AliasCB1", "TagCB1");
                table.Rows[1].Cells[0].Paragraphs[0].AddCheckBox(true, "AliasCB2", "TagCB2");
                table.Rows[1].Cells[1].Paragraphs[0].AddStructuredDocumentTag("Two", "AliasSdt2", "TagSdt2");
                document.Save(false);
                Assert.False(HasUnexpectedElements(document), "Document has unexpected elements. Order of elements matters!");
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var table = document.Tables[0];
                Assert.Equal(4, table.StructuredDocumentTags.Count);
                Assert.Equal(2, table.CheckBoxes.Count);
            }
        }
    }
}
