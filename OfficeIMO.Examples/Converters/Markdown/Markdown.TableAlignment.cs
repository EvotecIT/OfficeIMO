using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

namespace OfficeIMO.Examples.Markdown {
    internal static partial class Markdown {
        public static void Example_MarkdownTableAlignment(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "MarkdownTableAlignment.md");

            using var doc = WordDocument.Create();
            var table = doc.AddTable(2, 3);

            var left = table.Rows[0].Cells[0].Paragraphs[0];
            left.Text = "Left";
            left.ParagraphAlignment = JustificationValues.Left;

            var center = table.Rows[0].Cells[1].Paragraphs[0];
            center.Text = "Center";
            center.ParagraphAlignment = JustificationValues.Center;

            var right = table.Rows[0].Cells[2].Paragraphs[0];
            right.Text = "Right";
            right.ParagraphAlignment = JustificationValues.Right;

            table.Rows[1].Cells[0].Paragraphs[0].Text = "A";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "B";
            table.Rows[1].Cells[2].Paragraphs[0].Text = "C";

            doc.SaveAsMarkdown(filePath);
            Console.WriteLine($"✓ Created: {filePath}");

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
