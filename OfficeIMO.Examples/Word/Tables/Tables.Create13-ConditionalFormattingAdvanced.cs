using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Examples.Word {
    internal static partial class Tables {
        internal static void Example_ConditionalFormattingAdvanced(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating table with advanced conditional formatting");
            string filePath = System.IO.Path.Combine(folderPath, "ConditionalFormattingAdvanced.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(5, 2, WordTableStyle.PlainTable1);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Name";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Status";

                table.Rows[1].Cells[0].Paragraphs[0].Text = "Task1";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "Done";

                table.Rows[2].Cells[0].Paragraphs[0].Text = "Task2";
                table.Rows[2].Cells[1].Paragraphs[0].Text = "Pending";

                table.Rows[3].Cells[0].Paragraphs[0].Text = "Task3";
                table.Rows[3].Cells[1].Paragraphs[0].Text = "Skipped";

                table.Rows[4].Cells[0].Paragraphs[0].Text = "Task4";
                table.Rows[4].Cells[1].Paragraphs[0].Text = "Done";

                var builder = table.BeginConditionalFormatting();

                // rule 1: mark done tasks green else red and format text
                builder.AddRule(
                    "Status",
                    "Done",
                    TextMatchType.Equals,
                    Color.LightGreen,
                    Color.Black,
                    Color.LightPink,
                    Color.Black,
                    highlightColumns: new[] { "Name" },
                    matchTextFormat: p => p.SetBold(),
                    noMatchTextFormat: p => p.SetUnderline(UnderlineValues.Single));

                // rule 2: highlight pending tasks using a color object
                builder.AddRule(
                    "Status",
                    "Pending",
                    TextMatchType.Equals,
                    Color.Yellow,
                    null,
                    highlightColumns: new[] { "Name" },
                    matchTextFormat: p => p.SetItalic());

                // rule 3: highlight row when task is done and name starts with Task4
                builder.AddRule(
                    new[] {
                        ("Status", "Done", TextMatchType.Equals),
                        ("Name", "Task4", TextMatchType.StartsWith)
                    },
                    matchAll: true,
                    Color.LightSkyBlue,
                    highlightColumns: new[] { "Name" },
                    matchTextFormat: p => {
                        p.SetBold();
                        p.SetUnderline(UnderlineValues.Single);
                    });

                builder.Apply();

                document.Save(openWord);
            }
        }
    }
}
