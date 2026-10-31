using System.Linq;
using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Examples.Word {
    internal static partial class AdvancedDocument {

        /// <summary>
        /// Generates a second advanced document demonstrating additional scenarios.
        /// </summary>
        /// <param name="folderPath">Destination folder for the document.</param>
        /// <param name="openWord">Whether to open Word after creation.</param>
        public static void Example_AdvancedWord2(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating advanced document");
            string filePath = System.IO.Path.Combine(folderPath, "AdvancedDocument2.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var table = document.AddTable(1, 3);
                table.Alignment = WordTableAlignment.Right;
                table.Width = WordTableWidth.Percent(100);
                table.ColumnWidth = new List<WordTableWidth?> { WordTableWidth.Length(global::OfficeIMO.OfficeLength.Points(75)), WordTableWidth.Length(global::OfficeIMO.OfficeLength.Points(100)), WordTableWidth.Length(global::OfficeIMO.OfficeLength.Points(75)) };

                // add on the left
                table.Rows[0].Cells[0].Paragraphs[0].Text = "To: ";
                table.Rows[0].Cells[0].Paragraphs[0].LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
                table.Rows[0].Cells[0].Paragraphs[0].LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero;

                table.Rows[0].Cells[0].Paragraphs[0].AddParagraph().AddText("Customer Company");
                table.Rows[0].Cells[0].Paragraphs[1].LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
                table.Rows[0].Cells[0].Paragraphs[1].LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero;

                table.Rows[0].Cells[0].Paragraphs[1].AddParagraph().AddText("40-308, Mikołów");
                table.Rows[0].Cells[0].Paragraphs[2].LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
                table.Rows[0].Cells[0].Paragraphs[2].LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero;

                table.Rows[0].Cells[0].Paragraphs[2].AddParagraph().AddText("Poland");
                table.Rows[0].Cells[0].Paragraphs[3].LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
                table.Rows[0].Cells[0].Paragraphs[3].LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero;

                // add on the right
                table.Rows[0].Cells[2].Paragraphs[0].Text = "Evotec Services sp. z o.o.";
                table.Rows[0].Cells[2].Paragraphs[0].LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
                table.Rows[0].Cells[2].Paragraphs[0].LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero;

                table.Rows[0].Cells[2].Paragraphs[0].AddParagraph().AddText("ul. Drozdów 6");
                table.Rows[0].Cells[2].Paragraphs[1].LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
                table.Rows[0].Cells[2].Paragraphs[1].LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero;

                table.Rows[0].Cells[2].Paragraphs[1].AddParagraph().AddText("40-308, Mikołów");
                table.Rows[0].Cells[2].Paragraphs[2].LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
                table.Rows[0].Cells[2].Paragraphs[2].LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero;

                table.Rows[0].Cells[2].Paragraphs[2].AddParagraph().AddText("Poland");
                table.Rows[0].Cells[2].Paragraphs[3].LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
                table.Rows[0].Cells[2].Paragraphs[3].LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero;

                // lets hide the table visibility
                table.Style = WordTableStyle.TableNormal;

                // lets add some empty space
                document.AddParagraph();
                document.AddParagraph();

                // here's alternative way to build above
                var paragraph1 = document.AddParagraph("Evotec Services sp. z o.o.");
                paragraph1.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
                paragraph1.ParagraphAlignment = WordParagraphAlignment.Right;
                var paragraph2 = document.AddParagraph("ul. Drozdów 6");
                paragraph2.LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero;
                paragraph2.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
                paragraph2.ParagraphAlignment = WordParagraphAlignment.Right;
                var paragraph3 = document.AddParagraph("40-308, Mikołów");
                paragraph3.LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero;
                paragraph3.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
                paragraph3.ParagraphAlignment = WordParagraphAlignment.Right;
                var paragraph4 = document.AddParagraph("Poland");
                paragraph4.LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero;
                paragraph4.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
                paragraph4.ParagraphAlignment = WordParagraphAlignment.Right;

                document.Save();
                if (openWord) document.OpenInApplication();
            }
        }
    }
}
