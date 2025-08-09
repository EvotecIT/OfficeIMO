using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

namespace OfficeIMO.Examples.Markdown {
    internal static partial class Markdown {
        public static void Example_MarkdownNestedFormatting(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "MarkdownNestedFormatting.docx");
            string markdown = "Text ~~**bold strike**~~ and ***bold italic***.";
            var doc = markdown.LoadFromMarkdown(new MarkdownToWordOptions());
            doc.Save(filePath);
            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
