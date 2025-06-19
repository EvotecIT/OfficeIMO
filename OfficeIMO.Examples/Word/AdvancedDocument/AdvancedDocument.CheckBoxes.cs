using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class AdvancedDocument {
        public static void Example_CheckBoxesAdvanced(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating advanced check box document");
            string filePath = System.IO.Path.Combine(folderPath, "AdvancedCheckBoxes.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Options list:");

                var table = document.AddTable(3, 2, WordTableStyle.TableGrid);
                table.Rows[0].Cells[0].Paragraphs[0].AddCheckBox(true);
                table.Rows[0].Cells[1].Paragraphs[0].Text = "First option";

                table.Rows[1].Cells[0].Paragraphs[0].AddCheckBox(false);
                table.Rows[1].Cells[1].Paragraphs[0].AddHyperLink("Second link", new Uri("https://evotec.xyz"), true, "Hyperlink example");

                table.Rows[2].Cells[0].Paragraphs[0].AddCheckBox(true);
                table.Rows[2].Cells[1].Paragraphs[0].AddField(WordFieldType.Date);

                document.AddParagraph("Tasks:");
                var list = document.AddList(WordListStyle.Bulleted);
                list.AddItem("Task 1").AddCheckBox();
                list.AddItem("Task 2").AddCheckBox(true);

                document.AddHeadersAndFooters();

                document.Save(openWord);
            }
        }
    }
}
