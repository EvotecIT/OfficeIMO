using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;

internal static partial class Protect {

    internal static void Example_ProtectAlwaysReadOnly(string folderPath, bool openWord) {
        Console.WriteLine("[*] Creating standard document with protection 'Always Read Only'");
        string filePath = System.IO.Path.Combine(folderPath, "Basic Document with always read only protection.docx");
        using (WordDocument document = WordDocument.Create(filePath)) {
            var paragraph = document.AddParagraph("Basic paragraph - Page 4");
            paragraph.ParagraphAlignment = JustificationValues.Center;
            paragraph.Color = SixLabors.ImageSharp.Color.Blue;
            paragraph.AddText(" This is continuation in the same line");

            Console.WriteLine("Always read only: " + document.Settings.ReadOnlyRecommended);

            document.Settings.ReadOnlyRecommended = true;

            Console.WriteLine("Always read only: " + document.Settings.ReadOnlyRecommended);

            document.Settings.ReadOnlyRecommended = null;

            //document.Settings.ReadOnlyPassword = "Test123";

            document.Save(openWord);
        }
    }
}
