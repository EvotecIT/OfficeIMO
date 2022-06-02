using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Examples.Word {
    internal static partial class Tables {
        internal static void Example_TableBorders(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with all table styles");
            string filePath = System.IO.Path.Combine(folderPath, "Document with Table Styles.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Table With Borders");
                paragraph.ParagraphAlignment = JustificationValues.Center;

                WordTable wordTable = document.AddTable(4, 4, WordTableStyle.TableNormal);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
                wordTable.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";

                Console.WriteLine("Border Left: " + wordTable.Rows[1].Cells[0].Borders.LeftStyle);

                wordTable.Rows[1].Cells[1].Borders.LeftStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.LeftColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();

                wordTable.Rows[2].Cells[1].Borders.LeftStyle = BorderValues.Double;
                wordTable.Rows[2].Cells[1].Borders.LeftColor = Color.Gold;
                wordTable.Rows[2].Cells[1].Borders.LeftSize = 24;

                Console.WriteLine("Border Left Style: " + wordTable.Rows[1].Cells[1].Borders.LeftStyle);
                Console.WriteLine("Border Left Color: " + wordTable.Rows[1].Cells[1].Borders.LeftColorHex);

                document.Save(openWord);
            }
        }
    }
}
