using OfficeIMO.Word;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word {
    internal static partial class Lists {
        internal static void Example_DetectListStyles(string folderPath, bool openWord) {
            Console.WriteLine("[*] Detecting list style for paragraphs");
            string filePath = Path.Combine(folderPath, "ListStyleDetection.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var bulletList = document.AddList(WordListStyle.Bulleted);
                bulletList.AddItem("Bullet item");

                var numberedList = document.AddList(WordListStyle.Headings111);
                numberedList.AddItem("Numbered item");

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                foreach (var paragraph in document.Paragraphs) {
                    if (paragraph.IsListItem) {
                        Console.WriteLine($"{paragraph.Text} -> {paragraph.GetListStyle()}");
                    }
                }
                document.Save(openWord);
            }
        }
    }
}
