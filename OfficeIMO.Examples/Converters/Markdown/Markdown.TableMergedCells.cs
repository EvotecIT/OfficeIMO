using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using System;
using System.IO;

namespace OfficeIMO.Examples.Markdown {
    internal static partial class Markdown {
        public static void Example_MarkdownTableMergedCells(string folderPath, bool openWord) {
            string markdown = @"+---+---+---+
| AAAAA | B |
+ AAAAA +---+
| AAAAA | C |
+---+---+---+
| D | E | F |
+---+---+---+
";
            var doc = markdown.LoadFromMarkdown(new MarkdownToWordOptions());
            string filePath = Path.Combine(folderPath, "MarkdownTableMergedCells.docx");
            doc.Save(filePath);
            Console.WriteLine($"\u2713 Created: {filePath}");
            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}