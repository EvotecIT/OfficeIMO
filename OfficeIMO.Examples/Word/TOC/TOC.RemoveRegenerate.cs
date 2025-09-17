using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class TOC {
        internal static void Example_RemoveRegenerateTOC(string folderPath, bool openWord) {
            Console.WriteLine("[*] Removing and regenerating TOC");
            string filePath = System.IO.Path.Combine(folderPath, "DocumentTOCRemoveRegenerate.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var toc = document.AddTableOfContent();
                document.AddParagraph("Heading 1").Style = WordParagraphStyles.Heading1;
                toc.Remove();
                document.RegenerateTableOfContent();
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.Save(openWord);
            }
        }
    }
}
