using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Examples.Word {
    internal static partial class AdvancedDocument {

        public static void Example_AdvancedWord2(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating advanced document");
            string filePath = System.IO.Path.Combine(folderPath, "AdvancedDocument2.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var table = document.AddTable(1, 3);
                table.Alignment = TableRowAlignmentValues.Right;
                table.Width = 5000;
                table.WidthType = TableWidthUnitValues.Pct;
                table.ColumnWidth = new List<int>() { 1500, 2000, 1500 };

                // add on the left
                table.Rows[0].Cells[0].Paragraphs[0].Text = "To: ";
                table.Rows[0].Cells[0].Paragraphs[0].LineSpacingAfter = 0;
                table.Rows[0].Cells[0].Paragraphs[0].LineSpacingBefore = 0;

                table.Rows[0].Cells[0].Paragraphs[0].AddParagraph().AddText("Customer Company");
                table.Rows[0].Cells[0].Paragraphs[1].LineSpacingAfter = 0;
                table.Rows[0].Cells[0].Paragraphs[1].LineSpacingBefore = 0;

                table.Rows[0].Cells[0].Paragraphs[1].AddParagraph().AddText("40-308, Mikołów");
                table.Rows[0].Cells[0].Paragraphs[2].LineSpacingAfter = 0;
                table.Rows[0].Cells[0].Paragraphs[2].LineSpacingBefore = 0;

                table.Rows[0].Cells[0].Paragraphs[2].AddParagraph().AddText("Poland");
                table.Rows[0].Cells[0].Paragraphs[3].LineSpacingAfter = 0;
                table.Rows[0].Cells[0].Paragraphs[3].LineSpacingBefore = 0;

                // add on the right
                table.Rows[0].Cells[2].Paragraphs[0].Text = "Evotec Services sp. z o.o.";
                table.Rows[0].Cells[2].Paragraphs[0].LineSpacingAfter = 0;
                table.Rows[0].Cells[2].Paragraphs[0].LineSpacingBefore = 0;

                table.Rows[0].Cells[2].Paragraphs[0].AddParagraph().AddText("ul. Drozdów 6");
                table.Rows[0].Cells[2].Paragraphs[1].LineSpacingAfter = 0;
                table.Rows[0].Cells[2].Paragraphs[1].LineSpacingBefore = 0;

                table.Rows[0].Cells[2].Paragraphs[1].AddParagraph().AddText("40-308, Mikołów");
                table.Rows[0].Cells[2].Paragraphs[2].LineSpacingAfter = 0;
                table.Rows[0].Cells[2].Paragraphs[2].LineSpacingBefore = 0;

                table.Rows[0].Cells[2].Paragraphs[2].AddParagraph().AddText("Poland");
                table.Rows[0].Cells[2].Paragraphs[3].LineSpacingAfter = 0;
                table.Rows[0].Cells[2].Paragraphs[3].LineSpacingBefore = 0;

                // lets hide the table visibility
                table.Style = WordTableStyle.TableNormal;

                // lets add some empty space
                document.AddParagraph();
                document.AddParagraph();

                // here's alternative way to build above
                var paragraph1 = document.AddParagraph("Evotec Services sp. z o.o.");
                paragraph1.LineSpacingAfter = 0;
                paragraph1.ParagraphAlignment = JustificationValues.Right;
                var paragraph2 = document.AddParagraph("ul. Drozdów 6");
                paragraph2.LineSpacingBefore = 0;
                paragraph2.LineSpacingAfter = 0;
                paragraph2.ParagraphAlignment = JustificationValues.Right;
                var paragraph3 = document.AddParagraph("40-308, Mikołów");
                paragraph3.LineSpacingBefore = 0;
                paragraph3.LineSpacingAfter = 0;
                paragraph3.ParagraphAlignment = JustificationValues.Right;
                var paragraph4 = document.AddParagraph("Poland");
                paragraph4.LineSpacingBefore = 0;
                paragraph4.LineSpacingAfter = 0;
                paragraph4.ParagraphAlignment = JustificationValues.Right;

                document.Save(openWord);
            }
        }
    }
}
