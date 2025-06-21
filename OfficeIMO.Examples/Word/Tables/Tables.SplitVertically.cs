using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Tables {
        internal static void Example_SplitVertically(string folderPath, bool openWord) {
            Console.WriteLine("[*] Tables - split vertically merged cells");
            string filePath = System.IO.Path.Combine(folderPath, "TablesSplitVertically.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(4, 2, WordTableStyle.PlainTable1);

                table.Rows[0].Cells[0].Paragraphs[0].Text = "Row 1";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "Row 2";
                table.Rows[2].Cells[0].Paragraphs[0].Text = "Row 3";
                table.Rows[3].Cells[0].Paragraphs[0].Text = "Row 4";

                table.Rows[0].Cells[0].MergeVertically(2, true);
                table.Rows[0].Cells[0].SplitVertically(2);

                document.Save(openWord);
            }
        }
    }
}
