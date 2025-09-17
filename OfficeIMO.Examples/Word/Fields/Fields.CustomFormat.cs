using System;
using OfficeIMO.Examples.Utils;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Fields {
        internal static void Example_CustomFormattedDateField(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with custom formatted date field");
            string filePath = System.IO.Path.Combine(folderPath, "CustomFormattedDate.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Today is: ").AddField(WordFieldType.Date, customFormat: "dddd, MMMM dd, yyyy");
                document.Save(openWord);
            }
        }

        internal static void Example_CustomFormattedTimeField(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with custom formatted time field");
            string filePath = System.IO.Path.Combine(folderPath, "CustomFormattedTime.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Current time: ").AddField(WordFieldType.Time, customFormat: "HH:mm:ss");
                document.Save(openWord);
            }
        }

        internal static void Example_CustomFormattedHeaderDate(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with custom formatted date in header");
            string filePath = System.IO.Path.Combine(folderPath, "CustomFormattedHeaderDate.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();
                var headers = Guard.NotNull(document.Header, "Document headers must exist after enabling headers.");
                var defaultHeader = Guard.NotNull(headers.Default, "Default header must exist after enabling headers.");
                defaultHeader.AddField(WordFieldType.Date, customFormat: "yyyy-MM-dd", advanced: true);
                document.AddParagraph("Body paragraph");
                document.Save(openWord);
            }
        }
    }
}

