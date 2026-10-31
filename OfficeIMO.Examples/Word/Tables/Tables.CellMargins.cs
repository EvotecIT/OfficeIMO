using OfficeIMO.Word;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word {
    internal static partial class Tables {
        internal static void Example_TableCellMargins(string folderPath, bool openWord) {
            Console.WriteLine("[*] Table cell custom margins");
            string filePath = Path.Combine(folderPath, "TableCellMargins.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(2, 2);
                table.StyleDetails!.MarginDefaultRight = global::OfficeIMO.OfficeLength.Points(5);

                var cell1 = table.Rows[0].Cells[1];
                cell1.AddParagraph("Right margin 200 twips");
                cell1.MarginRight = global::OfficeIMO.OfficeLength.Points(10D);

                var cell2 = table.Rows[1].Cells[0];
                cell2.AddParagraph("Top margin 0.5 cm");
                cell2.MarginTop = global::OfficeIMO.OfficeLength.Centimeters(0.5D);

                document.Save();
                if (openWord) document.OpenInApplication();
            }
        }
    }
}
