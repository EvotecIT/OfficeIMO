using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Tables {
        internal static void Example_SplitHorizontally(string folderPath, bool openWord) {
            Console.WriteLine("[*] Tables - split horizontally merged cells");
            string filePath = System.IO.Path.Combine(folderPath, "TablesSplitHorizontally.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(2, 4, WordTableStyle.PlainTable1);

                table.Rows[0].Cells[0].Paragraphs[0].Text = "Col 1";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Col 2";
                table.Rows[0].Cells[2].Paragraphs[0].Text = "Col 3";
                table.Rows[0].Cells[3].Paragraphs[0].Text = "Col 4";

                table.Rows[0].Cells[1].MergeHorizontally(2, true);
                Console.WriteLine($"Merged: {table.Rows[0].Cells[1].HasHorizontalMerge}");
                table.Rows[0].Cells[1].SplitHorizontally(2);
                Console.WriteLine($"Merged after split: {table.Rows[0].Cells[1].HasHorizontalMerge}");

                document.Save(openWord);
            }
        }
    }
}
