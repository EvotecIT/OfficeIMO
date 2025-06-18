using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

internal static partial class Protect {

    internal static void Example_TestReadOnlyRecommendations(string folderPath, bool openWord) {
        Console.WriteLine("[*] Testing different read-only recommendation approaches");

        // Test 1: Just recommendation flag (no password)
        string filePath1 = System.IO.Path.Combine(folderPath, "Test1_RecommendationOnly.docx");
        using (WordDocument document = WordDocument.Create(filePath1)) {
            var paragraph = document.AddParagraph("Test 1: RECOMMENDATION ONLY (no password)");
            paragraph.ParagraphAlignment = JustificationValues.Center;
            paragraph.Color = SixLabors.ImageSharp.Color.Green;
            paragraph.AddText(" - Should show 'open as read-only?' dialog");

            // Only set recommendation, no password
            document.Settings.ReadOnlyRecommended = true;

            Console.WriteLine("Created Test 1: Recommendation only");
            document.Save(false);
        }

        // Test 2: Password only (no recommendation flag)
        string filePath2 = System.IO.Path.Combine(folderPath, "Test2_PasswordOnly.docx");
        using (WordDocument document = WordDocument.Create(filePath2)) {
            var paragraph = document.AddParagraph("Test 2: PASSWORD ONLY (no recommendation flag)");
            paragraph.ParagraphAlignment = JustificationValues.Center;
            paragraph.Color = SixLabors.ImageSharp.Color.Orange;
            paragraph.AddText(" - Should behave differently");

            // Only set password, no recommendation flag
            document.Settings.ReadOnlyPassword = "Test123";

            Console.WriteLine("Created Test 2: Password only");
            document.Save(false);
        }

        // Test 3: Both password and recommendation
        string filePath3 = System.IO.Path.Combine(folderPath, "Test3_PasswordAndRecommendation.docx");
        using (WordDocument document = WordDocument.Create(filePath3)) {
            var paragraph = document.AddParagraph("Test 3: BOTH PASSWORD AND RECOMMENDATION");
            paragraph.ParagraphAlignment = JustificationValues.Center;
            paragraph.Color = SixLabors.ImageSharp.Color.Purple;
            paragraph.AddText(" - Current behavior");

            // Set both password and recommendation
            document.Settings.ReadOnlyPassword = "Test123";
            document.Settings.ReadOnlyRecommended = true;

            Console.WriteLine("Created Test 3: Both password and recommendation");
            document.Save(false);
        }

        Console.WriteLine("\nTest all three documents in Word to see which behavior is correct!");

        if (openWord) {
            using (var document = WordDocument.Load(filePath1)) {
                document.Open(true);
            }
            using (var document = WordDocument.Load(filePath2)) {
                document.Open(true);
            }
            using (var document = WordDocument.Load(filePath3)) {
                document.Open(true);
            }
        }
    }
}
