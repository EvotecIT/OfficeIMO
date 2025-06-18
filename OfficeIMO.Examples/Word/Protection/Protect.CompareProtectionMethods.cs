using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

internal static partial class Protect {

    internal static void Example_ProtectDocumentPasswordVsReadOnlyRecommended(string folderPath, bool openWord) {
        Console.WriteLine("[*] Creating documents to demonstrate protection differences");

        // Method 1: Document Protection (enforced with password)
        string filePath1 = System.IO.Path.Combine(folderPath, "Document with ENFORCED password protection.docx");
        using (WordDocument document = WordDocument.Create(filePath1)) {
            var paragraph = document.AddParagraph("This document uses ENFORCED protection");
            paragraph.ParagraphAlignment = JustificationValues.Center;
            paragraph.Color = SixLabors.ImageSharp.Color.Red;
            paragraph.AddText(" - You MUST enter password 'Test123' to modify this document");

            // This creates ENFORCED document protection that requires password
            document.Settings.ProtectionPassword = "Test123";
            document.Settings.ProtectionType = DocumentProtectionValues.ReadOnly;

            Console.WriteLine("Created document with ENFORCED protection:");
            Console.WriteLine("- ProtectionType: " + document.Settings.ProtectionType);
            Console.WriteLine("- Password required to modify: Yes");

            document.Save(false);
        }

        // Method 2: Read-Only Recommendation (optional, user can ignore)
        string filePath2 = System.IO.Path.Combine(folderPath, "Document with read-only RECOMMENDATION.docx");
        using (WordDocument document = WordDocument.Create(filePath2)) {
            var paragraph = document.AddParagraph("This document only RECOMMENDS read-only");
            paragraph.ParagraphAlignment = JustificationValues.Center;
            paragraph.Color = SixLabors.ImageSharp.Color.Blue;
            paragraph.AddText(" - You can ignore this recommendation or enter password 'Test123'");            // This only recommends read-only mode (user can choose to ignore)
            document.Settings.ReadOnlyPassword = "Test123";
            document.Settings.ReadOnlyRecommended = true;

            Console.WriteLine("\nCreated document with read-only RECOMMENDATION:");
            Console.WriteLine("- ReadOnlyRecommended: " + document.Settings.ReadOnlyRecommended);
            Console.WriteLine("- Password to bypass recommendation: Optional");

            document.Save(false);
        }
        Console.WriteLine("\nKey Differences:");
        Console.WriteLine("1. ENFORCED Protection (ProtectionPassword): Password 'Test123' IS REQUIRED to modify");
        Console.WriteLine("2. READ-ONLY Recommendation (ReadOnlyPassword): Password 'Test123' is optional, user can ignore");
        Console.WriteLine("\nOpen both documents in Word to see the difference!");

        if (openWord) {
            // Open the first document to demonstrate enforced protection
            using (var document = WordDocument.Load(filePath1)) {
                document.Open(true);
            }
            using (var document = WordDocument.Load(filePath2)) {
                document.Open(true);
            }
        }
    }
}
