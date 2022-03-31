using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Tables {
        internal static void Example_AllTables(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with all table styles");
            string filePath = System.IO.Path.Combine(folderPath, "Document with Table Styles.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                //var listOfTablesStyles = Enum.GetValues(typeof(WordTableStyle)).Cast<WordTableStyle>();
                var listOfTablesStyles = (WordTableStyle[])Enum.GetValues(typeof(WordTableStyle));
                foreach (var tableStyle in listOfTablesStyles) {
                    var paragraph = document.AddParagraph(tableStyle.ToString());
                    paragraph.ParagraphAlignment = JustificationValues.Center;

                    WordTable wordTable = document.AddTable(4, 4, tableStyle);
                    wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                    wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                    wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
                    wordTable.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";
                }

                Console.WriteLine("+ Tables count: " + document.Tables.Count);

                document.Save(openWord);
            }
        }
    }
}
