using System;
using System.IO;
using OfficeIMO.Examples.Utils;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class DatePickers {
        internal static void Example_AdvancedDatePicker(string folderPath, bool openWord) {
            Console.WriteLine("[*] Updating date picker in existing document");
            string filePath = Path.Combine(folderPath, "DocumentWithDatePicker.docx");
            using (WordDocument document = WordDocument.Load(filePath)) {
                var picker = Guard.NotNull(document.GetDatePickerByTag("DateTag"), "Date picker with tag 'DateTag' was not found.");
                Console.WriteLine($"Current date: {picker.Date}");
                picker.Date = DateTime.Today.AddDays(1);
                document.Save(openWord);
            }
        }
    }
}
