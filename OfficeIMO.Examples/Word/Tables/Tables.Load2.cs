using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Tables {
        internal static void Example_BasicTablesLoad2(string templatesPath, string folderPath, bool openWord) {
            Console.WriteLine("[*] Loading standard document with tables created in Word");
            string filePath = System.IO.Path.Combine(templatesPath, "DocumentWithTables.docx");
            string filePath2 = System.IO.Path.Combine(folderPath, "DocumentWithTablesChanged.docx");
            using (WordDocument document = WordDocument.Load(filePath)) {
                Console.WriteLine(document.Tables.Count);

                var table = document.Tables[0];
                Console.WriteLine("First table style: " + table.Style);

                table.Style = WordTableStyle.GridTable1LightAccent4;

                Console.WriteLine("First table style, after change: " + table.Style);

                WordTable wordTable = document.AddTable(3, 4, WordTableStyle.GridTable1LightAccent5);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";

                wordTable = document.AddTable(3, 4, WordTableStyle.GridTable1LightAccent6);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";

                document.SaveAs(filePath2, openWord);
            }
        }
    }
}
