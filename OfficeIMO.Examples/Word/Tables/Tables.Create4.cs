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
                var paragraph = document.AddParagraph("Lets add table with no borders at all, and then lets fix it with some random borders for given cells: ");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Bold = true;
                paragraph.Underline = UnderlineValues.DotDash;

                WordTable wordTable = document.AddTable(4, 4, WordTableStyle.TableNormal);
                wordTable.RepeatHeaderRowAtTheTopOfEachPage = true;
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
                wordTable.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";

                Console.WriteLine("Border Left Style: " + wordTable.Rows[1].Cells[0].Borders.LeftStyle);

                wordTable.Rows[1].Cells[1].Borders.LeftStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.LeftColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();

                wordTable.Rows[2].Cells[1].Borders.LeftStyle = BorderValues.Double;
                wordTable.Rows[2].Cells[1].Borders.LeftColor = Color.Gold;
                wordTable.Rows[2].Cells[1].Borders.LeftSize = 24;

                Console.WriteLine("Border Left Style: " + wordTable.Rows[1].Cells[1].Borders.LeftStyle);
                Console.WriteLine("Border Left Color: " + wordTable.Rows[1].Cells[1].Borders.LeftColorHex);

                wordTable.Rows[2].Cells[1].Borders.TopLeftToBottomRightColor = Color.Aqua;
                wordTable.Rows[2].Cells[1].Borders.TopLeftToBottomRightSize = 24;
                wordTable.Rows[2].Cells[1].Borders.TopLeftToBottomRightStyle = BorderValues.Dotted;

                document.AddParagraph();
                document.AddHorizontalLine();
                paragraph = document.AddParagraph("Lets create new table with applied built-in style:");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Bold = true;
                paragraph.Underline = UnderlineValues.DotDash;


                var wordTable1 = document.AddTable(4, 4, WordTableStyle.GridTable5DarkAccent2);
                wordTable1.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable1.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable1.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
                wordTable1.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";

                document.AddParagraph();
                document.AddHorizontalLine();
                paragraph = document.AddParagraph("Lets create new table with default style, but lets fix it right after creating:");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Bold = true;
                paragraph.Underline = UnderlineValues.DotDash;

                var wordTable2 = document.AddTable(4, 4);
                wordTable2.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable2.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable2.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
                wordTable2.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";

                wordTable2.Style = WordTableStyle.GridTable7Colorful;


                wordTable.AddRow(5);

                document.Save(openWord);
            }
        }
    }
}
