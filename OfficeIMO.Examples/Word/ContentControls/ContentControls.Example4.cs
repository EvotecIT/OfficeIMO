using OfficeIMO.Word;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word {
    internal static partial class ContentControls {
        internal static void Example_ContentControlsInTable(string folderPath, bool openWord) {
            Console.WriteLine("[*] Content controls inside a table");
            string filePath = Path.Combine(folderPath, "DocumentContentControlsInTable.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var table = document.AddTable(2, 2);
                table.Rows[0].Cells[0].Paragraphs[0].AddStructuredDocumentTag(alias: "Alias1", text: "One", tag: "Tag1");
                var cb = table.Rows[1].Cells[1].Paragraphs[0].AddCheckBox(false, "CheckAlias", "CheckTag");
                document.Save(openWord);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var table = document.Tables[0];
                foreach (var control in table.StructuredDocumentTags) {
                    Console.WriteLine($"SDT {control.Tag}: {control.Text}");
                }
                foreach (var checkBox in table.CheckBoxes) {
                    Console.WriteLine($"CheckBox {checkBox.Tag}: alias={checkBox.Alias}");
                }
                document.Save(openWord);
            }
        }
    }
}
