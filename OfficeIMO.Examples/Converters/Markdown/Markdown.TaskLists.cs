using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

namespace OfficeIMO.Examples.Markdown {
    internal static partial class Markdown {
        public static void Example_MarkdownTaskLists(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "MarkdownTaskLists.docx");
            string markdown = "- [ ] Task 1\n- [x] Task 2\n  - [ ] Subtask";

            var doc = markdown.LoadFromMarkdown(new MarkdownToWordOptions());
            doc.Save(filePath);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
