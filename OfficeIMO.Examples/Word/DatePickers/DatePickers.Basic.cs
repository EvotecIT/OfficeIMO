using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class DatePickers {
        internal static void Example_BasicDatePicker(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with a date picker control");
            string filePath = Path.Combine(folderPath, "DocumentWithDatePicker.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Date: ");
                paragraph.AddDatePicker(DateTime.Today, "DateAlias", "DateTag");
                document.Save(openWord);
            }
        }
    }
}
