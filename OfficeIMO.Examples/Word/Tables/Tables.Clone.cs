using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Tables {
        internal static void Example_CloneTable(string folderPath, bool openWord) {
            Console.WriteLine("[*] Cloning table");
            string filePath = System.IO.Path.Combine(folderPath, "Document with Cloned Table.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(2, 2, WordTableStyle.GridTable1LightAccent1);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "A1";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "A2";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "B1";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "B2";

                WordTable cloned = table.Clone();
                cloned.Rows[0].Cells[0].Paragraphs[0].Bold = true;

                document.Save(openWord);
            }
        }
    }
}
