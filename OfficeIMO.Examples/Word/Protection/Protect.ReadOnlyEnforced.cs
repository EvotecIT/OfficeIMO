using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word;

internal static partial class Protect {
    // Example: Enforced read-only protection (password required to edit)
    internal static void Example_ReadOnlyEnforced(string folderPath, bool openWord) {
        Console.WriteLine("[*] Creating document with enforced read-only protection (password required to edit)");
        string filePath = System.IO.Path.Combine(folderPath, "Basic Document with enforced read-only protection.docx");
        using (WordDocument document = WordDocument.Create(filePath)) {
            var paragraph = document.AddParagraph("This document is protected: password required to edit");
            paragraph.ParagraphAlignment = JustificationValues.Center;
            paragraph.Color = SixLabors.ImageSharp.Color.Red;
            paragraph.AddText(" - Password is 'Test123'");

            // This is the only way to enforce editing restrictions in modern Word
            document.Settings.ProtectionPassword = "Test123";
            document.Settings.ProtectionType = DocumentProtectionValues.ReadOnly;

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