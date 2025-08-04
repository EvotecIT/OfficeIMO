using OfficeIMO.Word;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word {
    internal static partial class Tables {
        internal static void Example_TableCellOptions(string folderPath, bool openWord) {
            Console.WriteLine("[*] Table cell WrapText and FitText options");
            string filePath = Path.Combine(folderPath, "TableCellOptions.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(2, 2);

                var cell1 = table.Rows[0].Cells[0];
                cell1.AddParagraph("Default behavior");

                var cell2 = table.Rows[0].Cells[1];
                cell2.AddParagraph("No wrap");
                cell2.WrapText = false;

                var cell3 = table.Rows[1].Cells[0];
                cell3.AddParagraph("Fit text");
                cell3.FitText = true;

                var cell4 = table.Rows[1].Cells[1];
                cell4.AddParagraph("No wrap & fit");
                cell4.WrapText = false;
                cell4.FitText = true;

                document.Save(openWord);
            }
        }
    }
}