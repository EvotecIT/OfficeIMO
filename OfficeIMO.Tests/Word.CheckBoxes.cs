using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_AddingCheckBoxes() {
            string filePath = Path.Combine(_directoryWithFiles, "CheckBoxDocument.docx");
            using (var document = WordDocument.Create(filePath)) {
                var p1 = document.AddParagraph("Option 1");
                p1.AddCheckBox(true);
                var p2 = document.AddParagraph("Option 2");
                p2.AddCheckBox();

                var table = document.AddTable(2, 2);
                table.Rows[0].Cells[0].Paragraphs[0].AddCheckBox();
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Table option";

                var list = document.AddList(WordListStyle.Bulleted);
                list.AddItem("Task 1").AddCheckBox();

                Assert.True(document.Paragraphs[0].IsCheckBox);
                Assert.True(document.Paragraphs[0].CheckBox.Checked);
                Assert.False(document.Paragraphs[1].CheckBox.Checked);

                document.Paragraphs[1].CheckBox.Checked = true;
                Assert.True(document.Paragraphs[1].CheckBox.Checked);

                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.Equal(5, document.Paragraphs.Count);
                Assert.All(document.Paragraphs, p => Assert.True(p.IsCheckBox || p.Text != null));
                Assert.True(document.Paragraphs[1].CheckBox.Checked);
            }
        }
    }
}
