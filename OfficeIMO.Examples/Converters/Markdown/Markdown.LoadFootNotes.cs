using System.IO;
using OfficeIMO.Word.Markdown;

namespace OfficeIMO.Examples.Markdown {
    internal static partial class Markdown {
        public static void Example_MarkdownLoadFootNotes(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "MarkdownLoadFootNotes.docx");
            string md = "Paragraph with footnote[^1].\n\n[^1]: Footnote text";
            using var document = md.LoadFromMarkdown();
            document.Save(filePath);
            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
