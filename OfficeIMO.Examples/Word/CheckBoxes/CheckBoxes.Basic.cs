using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class CheckBoxes {
        internal static void Example_BasicCheckBox(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with a checkbox control");
            string filePath = Path.Combine(folderPath, "DocumentWithCheckBox.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Accept terms:");
                var checkBox = paragraph.AddCheckBox(true, "Terms");
                Console.WriteLine($"Checkbox initially checked: {checkBox.IsChecked}");
                document.Save(openWord);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var checkBox = document.CheckBoxes[0];
                Console.WriteLine($"Loaded state: {checkBox.IsChecked}");
                checkBox.IsChecked = false;
                document.Save(openWord);
            }
        }
    }
}
