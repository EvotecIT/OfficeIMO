using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Fields {
        internal static void Example_FieldWithMultipleSwitches(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating field with multiple format switches");
            string filePath = System.IO.Path.Combine(folderPath, "FieldMultipleSwitchesExample.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph().AddField(WordFieldType.Page, WordFieldFormat.Caps);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var field = document.Fields[0];
                Console.WriteLine("Format switches: " + String.Join(", ", field.FieldFormat));
                document.Save(openWord);
            }
        }
    }
}
