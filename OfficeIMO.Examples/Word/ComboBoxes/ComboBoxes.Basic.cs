using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class ComboBoxes {
        internal static void Example_BasicComboBox(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with a combo box control");
            string filePath = Path.Combine(folderPath, "DocumentWithComboBox.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var items = new[] { "One", "Two", "Three" };
                document.AddParagraph("Choose:").AddComboBox(items, "Combo", "ComboTag", defaultValue: "Two");
                document.Save(openWord);
            }
        }
    }
}
