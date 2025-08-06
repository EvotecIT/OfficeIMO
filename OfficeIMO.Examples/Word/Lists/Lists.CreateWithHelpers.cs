using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Lists {
        internal static void Example_CreateListsWithHelpers(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with helper list methods");
            string filePath = Path.Combine(folderPath, "DocumentWithHelperLists.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var bullets = document.AddListBulleted();
                bullets.AddItem("Bullet 1");
                bullets.AddItem("Bullet 2");

                var numbers = document.AddListNumbered();
                numbers.AddItem("First");
                numbers.AddItem("Second");

                document.Save(openWord);
            }
        }
    }
}
