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
                Console.WriteLine($"Bullet list uses style: {bulletList.Style}");

                var numberedList = document.AddList(WordListStyle.Numbered);
                numberedList.AddItem("Numbered item");
                Console.WriteLine($"Numbered list uses style: {numberedList.Style}");

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                foreach (var list in document.Lists) {
                    Console.WriteLine($"Loaded list style: {list.Style}");
                }
                foreach (var paragraph in document.Paragraphs) {
                    if (paragraph.IsListItem) {
                        Console.WriteLine($"{paragraph.Text} -> {paragraph.ListStyle}");
                    }
                }
                document.Save(openWord);
            }
        }
    }
}
