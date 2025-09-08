using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word {
    internal static partial class Tables {
        internal static void Example_TableCustomWidthAndShading(string folderPath, bool openWord) {
            Console.WriteLine("[*] Table with custom column widths and shading");
            string filePath = Path.Combine(folderPath, "TableCustomWidthAndShading.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(2, 2);

                table.ColumnWidth = new List<int> { 1440, 2880 };
                table.ColumnWidthType = TableWidthUnitValues.Dxa;

                var cell1 = table.Rows[0].Cells[0];
                cell1.AddParagraph("Red");
                cell1.ShadingFillColorHex = "ff0000";

                var cell2 = table.Rows[0].Cells[1];
                cell2.AddParagraph("Blue");
                cell2.ShadingFillColorHex = "0000ff";

                document.Save(openWord);
            }
        }
    }
}

