using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Tables {
        internal static void Example_ConditionalFormattingValues(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating table with conditional formatting based on cell values");
            string filePath = System.IO.Path.Combine(folderPath, "ConditionalFormattingValues.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(3, 2, WordTableStyle.PlainTable1);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Name";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Status";

                table.Rows[1].Cells[0].Paragraphs[0].Text = "Task1";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "Done";
                table.Rows[2].Cells[0].Paragraphs[0].Text = "Task2";
                table.Rows[2].Cells[1].Paragraphs[0].Text = "Pending";

                table.ConditionalFormatting(
                    "Status",
                    "Done",
                    TextMatchType.Equals,
                    matchFillColorHex: "92d050",
                    noMatchFillColorHex: "ff0000",
                    matchTextFormat: p => p.SetBold(),
                    noMatchTextFormat: p => p.SetUnderline(UnderlineValues.Single));

                table.ConditionalFormatting(
                    new[] {
                        ("Status", "Done", TextMatchType.Equals),
                        ("Name", "Task1", TextMatchType.StartsWith)
                    },
                    matchAll: true,
                    matchFillColorHex: "92d050",
                    noMatchFillColorHex: "ff0000",
                    highlightColumns: new[] { "Name" },
                    matchTextFormat: p => p.SetBold(),
                    noMatchTextFormat: p => p.SetUnderline(UnderlineValues.Single));

                document.Save(openWord);
            }
        }
    }
}
