using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;

internal static partial class Protect {

    internal static void Example_ProtectFinalDocument(string folderPath, bool openWord) {
        Console.WriteLine("[*] Creating basic document with protection - Final Document");
        string filePath = System.IO.Path.Combine(folderPath, "Basic Document with setting Word to Final Document.docx");
        using (WordDocument document = WordDocument.Create(filePath)) {
            var paragraph = document.AddParagraph("Basic paragraph - Page 1");
            paragraph.ParagraphAlignment = JustificationValues.Center;
            paragraph.Color = SixLabors.ImageSharp.Color.Blue;
            paragraph.AddText(" This is continutation in the same line");
            paragraph.AddBreak(BreakValues.TextWrapping);

            Console.WriteLine("Final document: " + document.Settings.FinalDocument);

            document.Settings.FinalDocument = true;

            Console.WriteLine("Final document: " + document.Settings.FinalDocument);

            document.Settings.FinalDocument = false;

            document.Save(openWord);
        }
    }
}
