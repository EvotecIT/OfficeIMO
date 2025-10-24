using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Tables {
        internal static void Example_TablesParentAccess(string folderPath, bool openWord) {
            Console.WriteLine("[*] Accessing table row information from a cell");
            string filePath = System.IO.Path.Combine(folderPath, "Document with Table Parent Access.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(1, 2, WordTableStyle.TableGrid);
                table.Rows[0].Height = 480;

                WordTableCell firstCell = table.Rows[0].Cells[0];
                firstCell.Paragraphs[0].Text = "Row height:";

                int? heightFromCell = firstCell.Parent.Height;
                firstCell.AddParagraph(heightFromCell?.ToString() ?? "not set");

                Console.WriteLine($"Row height read from cell: {heightFromCell}");
                Console.WriteLine($"Cell belongs to table style: {firstCell.ParentTable.Style}");

                document.Save(openWord);
            }
        }
    }
}
