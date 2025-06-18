using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

internal static partial class Protect {

    internal static void Example_ReadOnlyRecommended(string folderPath, bool openWord) {
        Console.WriteLine("[*] Creating document with 'Read-Only Recommended' (suggestion, not enforced)");
        string filePath = System.IO.Path.Combine(folderPath, "Basic Document with read-only recommended.docx");
        using (WordDocument document = WordDocument.Create(filePath)) {
            var paragraph = document.AddParagraph("Basic paragraph - Page 4");
            paragraph.ParagraphAlignment = JustificationValues.Center;
            paragraph.Color = SixLabors.ImageSharp.Color.Blue;
            paragraph.AddText(" This is continuation in the same line");

            Console.WriteLine("ReadOnlyRecommended: " + document.Settings.ReadOnlyRecommended);

            document.Settings.ReadOnlyRecommended = true;

            Console.WriteLine("ReadOnlyRecommended: " + document.Settings.ReadOnlyRecommended);

            document.Save(false);
            var valid = document.ValidateDocument();
            if (valid.Count > 0) {
                Console.WriteLine("Document has validation errors:");
                foreach (var error in valid) {
                    Console.WriteLine(error.Id + ": " + error.Description);
                }
            } else {
                Console.WriteLine("Document is valid.");
            }

            document.Open(openWord);
        }
    }
}
