using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Examples.Word;

internal static partial class Protect {

    // Example: Mark document as Final (shows 'Mark as Final' banner in Word)
    internal static void Example_FinalDocument(string folderPath, bool openWord) {
        Console.WriteLine("[*] Creating basic document with 'Final Document' property");
        string filePath = System.IO.Path.Combine(folderPath, "Basic Document with setting Word to Final Document.docx");
        using (WordDocument document = WordDocument.Create(filePath)) {
            var paragraph = document.AddParagraph("Basic paragraph - Page 1");
            paragraph.ParagraphAlignment = JustificationValues.Center;
            paragraph.Color = OfficeIMO.Drawing.OfficeColor.Blue;
            paragraph.AddText(" This is continutation in the same line");
            paragraph.AddBreak(BreakValues.TextWrapping);

            Console.WriteLine("Final document: " + document.Settings.FinalDocument);

            document.Settings.FinalDocument = true;

            Console.WriteLine("Final document: " + document.Settings.FinalDocument);

            document.Save(openWord);
        }
    }
}
